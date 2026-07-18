"""
onboarding.py
=============
Microsoft 365 Group / Team / DL Membership Manager — with Owner Approval
Uses Microsoft Graph API. Nothing is ever added or removed without the
owner (your manager) approving it first.

──────────────────────────────────────────────────────────────────────────────
  HOW IT WORKS  (request → approve → execute)
──────────────────────────────────────────────────────────────────────────────
  Anyone can run:
    1) LIST    → read-only. Shows a user's current membership everywhere.
                 No approval needed — nothing is changed.
    2) ADD     → creates a PENDING request to add a user to chosen
                 resource(s). Nothing happens yet.
    3) REMOVE  → creates a PENDING request to remove a user's MEMBERSHIP
                 from chosen resource(s). The Group/Team/DL is NEVER
                 deleted — only the membership.
    4) MOVE TEAM → creates a linked pair of PENDING requests: REMOVE from
                 the old team + ADD to the new team, so a person moving
                 teams doesn't stay in the old one.

  Only the owner runs:
    5) APPROVE → lists every PENDING request with full detail, asks
                 "Approve? (y/n)" one at a time, and ONLY THEN calls
                 Microsoft Graph to actually add/remove the membership.
                 Denied requests are logged and closed, not executed.

  Pending requests live in pending_requests.json next to this script,
  so they survive between "someone submits" and "owner approves" runs.

──────────────────────────────────────────────────────────────────────────────
  SETUP — Azure AD App Registration Permissions required
──────────────────────────────────────────────────────────────────────────────
  Add these APPLICATION permissions, then click "Grant admin consent":
    User.Read.All               ← look up the employee by email
    Group.ReadWrite.All         ← read & modify group membership
    GroupMember.ReadWrite.All   ← add / remove group members
    TeamMember.ReadWrite.All    ← add / remove Team members
    Sites.FullControl.All       ← grant/revoke SharePoint permissions
    Directory.ReadWrite.All     ← general directory access
    Mail.Send                   ← OPTIONAL, only needed to email the owner
                                   a notification when a request is created

──────────────────────────────────────────────────────────────────────────────
  SETUP — Environment variables  (.env file or export in terminal)
──────────────────────────────────────────────────────────────────────────────
    MS_TENANT_ID      – Azure AD Tenant ID
    MS_CLIENT_ID      – App (client) ID of your Azure AD App Registration
    MS_CLIENT_SECRET  – Client Secret for that app
    OWNER_EMAIL       – Your manager's email. Used only to send a
                         notification when a new request is created.
                         (Optional — if not set, notification is skipped,
                         requests still queue up fine for Approve mode.)

Usage:
  pip install msal requests python-dotenv
  python onboarding.py
"""

import os
import sys
import json
import time
import uuid
import requests
from datetime import datetime, timezone
from dotenv import load_dotenv

load_dotenv()

try:
    import msal
except ImportError:
    print("[ERROR] msal is not installed. Run: pip install msal")
    sys.exit(1)


# ==============================================================================
#  !! EDIT ONLY THIS SECTION !!
#
#  HOW TO FIND OBJECT IDs
#    Groups / DLs  → Azure AD Portal → Groups → click the group → "Object ID"
#    Teams         → Azure AD Portal → Groups → filter by "Teams" → Object ID
#                    (A Team is backed by an M365 Group — same Object ID)
#    SharePoint    → Graph call: GET /v1.0/sites/{hostname}:/sites/{sitename}
# ==============================================================================

M365_GROUPS: dict[str, str] = {
    "Everyone":                 "00000000-0000-0000-0000-000000000001",
    "Microsoft 365 Main Group": "00000000-0000-0000-0000-000000000002",
    "Security Group - General": "00000000-0000-0000-0000-000000000003",
}

DISTRIBUTION_LISTS: dict[str, str] = {
    "All Staff DL":             "00000000-0000-0000-0000-000000000010",
    "Company Announcements DL": "00000000-0000-0000-0000-000000000011",
    "Hiring":                   "00000000-0000-0000-0000-000000000012",
}

MS_TEAMS: dict[str, str] = {
    "TA - Alpha Team":        "00000000-0000-0000-0000-000000000020",
    "TA - Beta Team":         "00000000-0000-0000-0000-000000000021",
    "TA - Gamma Team":        "00000000-0000-0000-0000-000000000022",
    "TL Team":                "00000000-0000-0000-0000-000000000024",
    "Software Engineer Team": "00000000-0000-0000-0000-000000000025",
}

# "Label": ("Graph Site ID", "role")   role: "read" | "write" | "fullControl"
SHAREPOINT_SITES: dict[str, tuple[str, str]] = {
    "Home SharePoint":        ("contoso.sharepoint.com,aaa-bbb,ccc-ddd", "read"),
    "Recruitment SharePoint": ("contoso.sharepoint.com,eee-fff,ggg-hhh", "read"),
}

# NOTE: the names/IDs above are placeholders — replace them with your
# real company groups (every "TA - <Team>", "Everyone", "Hiring", etc).
# Whatever you put in these four dictionaries is exactly what gets
# checked automatically for every employee, no extra setup needed.

# ==============================================================================
#  END OF CONFIGURATION
# ==============================================================================

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
PENDING_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pending_requests.json")

_token_cache: dict = {}

# All configured resources in one lookup table, tagged by type, so
# List/Add/Remove/Move can all iterate one structure.
# type: "group" | "dl" | "team" | "sharepoint"
def _all_resources() -> list[dict]:
    resources = []
    for label, rid in M365_GROUPS.items():
        resources.append({"type": "group", "label": label, "id": rid})
    for label, rid in DISTRIBUTION_LISTS.items():
        resources.append({"type": "dl", "label": label, "id": rid})
    for label, rid in MS_TEAMS.items():
        resources.append({"type": "team", "label": label, "id": rid})
    for label, (site_id, role) in SHAREPOINT_SITES.items():
        resources.append({"type": "sharepoint", "label": label, "id": site_id, "role": role})
    return resources


# ==============================================================================
# AUTHENTICATION
# ==============================================================================

def get_access_token() -> str:
    now = time.time()
    if _token_cache.get("expires_at", 0) > now + 60:
        return _token_cache["token"]

    tenant_id     = _require_env("MS_TENANT_ID")
    client_id     = _require_env("MS_CLIENT_ID")
    client_secret = _require_env("MS_CLIENT_SECRET")

    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
    )
    result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

    if "access_token" not in result:
        err = result.get("error_description") or result.get("error") or str(result)
        raise RuntimeError(f"Failed to acquire access token: {err}")

    _token_cache["token"]      = result["access_token"]
    _token_cache["expires_at"] = now + result.get("expires_in", 3600)
    return _token_cache["token"]


def _require_env(name: str) -> str:
    value = os.environ.get(name, "").strip()
    if not value:
        print(f"\n[ERROR] Environment variable '{name}' is not set.")
        print("        Add it to your .env file or export it in your terminal.")
        sys.exit(1)
    return value


def _headers() -> dict:
    return {"Authorization": f"Bearer {get_access_token()}", "Content-Type": "application/json"}


# ==============================================================================
# GRAPH API — HELPERS
# ==============================================================================

def graph_get(path: str, params: dict = None) -> dict:
    resp = requests.get(f"{GRAPH_BASE}{path}", headers=_headers(), params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def graph_get_paged(path: str) -> list:
    items, url = [], f"{GRAPH_BASE}{path}"
    while url:
        resp = requests.get(url, headers=_headers(), timeout=30)
        resp.raise_for_status()
        data = resp.json()
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def graph_post(path: str, payload: dict) -> dict:
    resp = requests.post(f"{GRAPH_BASE}{path}", headers=_headers(), json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json() if resp.content else {}


def graph_delete(path: str) -> None:
    resp = requests.delete(f"{GRAPH_BASE}{path}", headers=_headers(), timeout=30)
    resp.raise_for_status()


# ==============================================================================
# USER RESOLUTION
# ==============================================================================

def resolve_user(email: str) -> dict | None:
    try:
        return graph_get(f"/users/{email}", params={"$select": "id,displayName,mail,userPrincipalName"})
    except requests.HTTPError as exc:
        status = exc.response.status_code
        if status == 404:
            return None
        if status == 403:
            print("\n[ERROR] 403 Forbidden when looking up user.")
            print("        Your Azure AD app is missing the 'User.Read.All' permission.")
            sys.exit(1)
        raise


# ==============================================================================
# MEMBERSHIP CHECK HELPERS
# ==============================================================================

def is_group_member(user_id: str, group_id: str) -> bool:
    try:
        data = graph_get(f"/users/{user_id}/memberOf", params={"$filter": f"id eq '{group_id}'", "$select": "id"})
        return len(data.get("value", [])) > 0
    except requests.HTTPError as exc:
        if exc.response.status_code == 400:
            members = graph_get_paged(f"/groups/{group_id}/members")
            return any(m.get("id") == user_id for m in members)
        raise


def is_team_member(user_id: str, team_id: str) -> bool:
    members = graph_get_paged(f"/teams/{team_id}/members")
    return any(m.get("userId") == user_id for m in members)


def get_team_membership_id(user_id: str, team_id: str) -> str | None:
    members = graph_get_paged(f"/teams/{team_id}/members")
    for m in members:
        if m.get("userId") == user_id:
            return m.get("id")
    return None


def has_sharepoint_access(user_email: str, site_id: str) -> tuple[bool, str | None]:
    """Returns (has_access, permission_id) so remove can target the exact grant."""
    try:
        perms = graph_get_paged(f"/sites/{site_id}/permissions")
        for perm in perms:
            for identity in perm.get("grantedToIdentities") or []:
                if identity.get("user", {}).get("email", "").lower() == user_email.lower():
                    return True, perm.get("id")
        return False, None
    except requests.HTTPError:
        return False, None


def is_member(resource: dict, user_id: str, user_email: str) -> bool:
    if resource["type"] in ("group", "dl"):
        return is_group_member(user_id, resource["id"])
    if resource["type"] == "team":
        return is_team_member(user_id, resource["id"])
    if resource["type"] == "sharepoint":
        has_access, _ = has_sharepoint_access(user_email, resource["id"])
        return has_access
    return False


# ==============================================================================
# EXECUTE HELPERS  (only ever called from Approve mode, after a "yes")
# ==============================================================================

def add_to_group(user_id: str, group_id: str) -> None:
    graph_post(f"/groups/{group_id}/members/$ref", {"@odata.id": f"{GRAPH_BASE}/directoryObjects/{user_id}"})


def add_to_team(user_id: str, team_id: str) -> None:
    graph_post(f"/teams/{team_id}/members", {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": [],
        "user@odata.bind": f"{GRAPH_BASE}/users('{user_id}')",
    })


def grant_sharepoint(user_email: str, site_id: str, role: str) -> None:
    graph_post(f"/sites/{site_id}/permissions", {
        "roles": [role],
        "grantedToIdentities": [{"user": {"email": user_email}}],
    })


def remove_from_group(user_id: str, group_id: str) -> None:
    """Removes MEMBERSHIP only. The Group/DL itself is never deleted."""
    graph_delete(f"/groups/{group_id}/members/{user_id}/$ref")


def remove_from_team(user_id: str, team_id: str) -> None:
    """Removes MEMBERSHIP only. The Team itself is never deleted."""
    membership_id = get_team_membership_id(user_id, team_id)
    if membership_id is None:
        return
    graph_delete(f"/teams/{team_id}/members/{membership_id}")


def revoke_sharepoint(user_email: str, site_id: str) -> None:
    """Removes the permission grant only. The Site itself is never touched."""
    _, perm_id = has_sharepoint_access(user_email, site_id)
    if perm_id is None:
        return
    graph_delete(f"/sites/{site_id}/permissions/{perm_id}")


def execute_request(req: dict) -> None:
    """Actually performs the add/remove on Graph. Only called after approval."""
    rtype, action = req["resource_type"], req["action"]
    if action == "ADD":
        if rtype in ("group", "dl"):
            add_to_group(req["user_id"], req["resource_id"])
        elif rtype == "team":
            add_to_team(req["user_id"], req["resource_id"])
        elif rtype == "sharepoint":
            grant_sharepoint(req["employee_email"], req["resource_id"], req.get("role", "read"))
    else:  # REMOVE
        if rtype in ("group", "dl"):
            remove_from_group(req["user_id"], req["resource_id"])
        elif rtype == "team":
            remove_from_team(req["user_id"], req["resource_id"])
        elif rtype == "sharepoint":
            revoke_sharepoint(req["employee_email"], req["resource_id"])


# ==============================================================================
# OWNER NOTIFICATION  (best-effort — never blocks request creation)
# ==============================================================================

def notify_owner(req: dict) -> None:
    owner_email = os.environ.get("OWNER_EMAIL", "").strip()
    if not owner_email:
        print("  (No OWNER_EMAIL configured — skipping email notification.")
        print("   The owner can still see this in Approve mode.)")
        return
    subject = f"[Approval needed] {req['action']} — {req['employee_name']} — {req['resource_label']}"
    body = (
        f"A new membership request needs your approval.\n\n"
        f"Employee : {req['employee_name']} ({req['employee_email']})\n"
        f"Action   : {req['action']}\n"
        f"Resource : {req['resource_label']} ({req['resource_type']})\n"
        f"Reason   : {req.get('reason', '-')}\n"
        f"Request ID: {req['id']}\n\n"
        f"Run this script and choose 'Approve pending requests' to act on it."
    )
    try:
        graph_post("/me/sendMail", {
            "message": {
                "subject": subject,
                "body": {"contentType": "Text", "content": body},
                "toRecipients": [{"emailAddress": {"address": owner_email}}],
            }
        })
        print(f"  [MAILED] Notification sent to owner ({owner_email})")
    except requests.HTTPError as exc:
        print(f"  [WARN] Could not email owner ({exc}). Request is still queued for Approve mode.")


# ==============================================================================
# PENDING REQUEST QUEUE
# ==============================================================================

def load_pending() -> list[dict]:
    if not os.path.exists(PENDING_FILE):
        return []
    with open(PENDING_FILE, "r") as f:
        return json.load(f)


def save_pending(requests_list: list[dict]) -> None:
    with open(PENDING_FILE, "w") as f:
        json.dump(requests_list, f, indent=2)


def queue_request(action: str, resource: dict, user: dict, reason: str = "") -> dict:
    req = {
        "id": uuid.uuid4().hex[:8],
        "created_at": datetime.now(timezone.utc).isoformat(),
        "action": action,                      # "ADD" | "REMOVE"
        "resource_type": resource["type"],
        "resource_label": resource["label"],
        "resource_id": resource["id"],
        "role": resource.get("role"),
        "employee_email": user["mail"] or user["userPrincipalName"],
        "employee_name": user.get("displayName") or user["userPrincipalName"],
        "user_id": user["id"],
        "reason": reason,
        "status": "PENDING",
    }
    pending = load_pending()
    pending.append(req)
    save_pending(pending)
    print(f"  [QUEUED] {action} → {resource['label']}  (request id: {req['id']}, status: PENDING)")
    notify_owner(req)
    return req


# ==============================================================================
# PRINT HELPERS
# ==============================================================================

def _banner(text: str) -> None:
    print(f"\n{'─' * 62}\n  {text}\n{'─' * 62}")


def _safe_error(exc: requests.HTTPError) -> str:
    try:
        body = exc.response.json()
        return body.get("error", {}).get("message", exc.response.text[:200])
    except Exception:
        return exc.response.text[:200]


def _resource_menu(resources: list[dict]) -> list[dict]:
    for i, r in enumerate(resources, start=1):
        print(f"  {i:>2}) [{r['type'].upper():<10}] {r['label']}")
    print(f"  {len(resources) + 1:>2}) ALL of the above")
    raw = input("\n  Enter numbers separated by commas (or ALL number): ").strip()
    if raw == str(len(resources) + 1):
        return resources
    chosen = []
    for part in raw.split(","):
        part = part.strip()
        if part.isdigit() and 1 <= int(part) <= len(resources):
            chosen.append(resources[int(part) - 1])
    return chosen


# ==============================================================================
# MODE 1 — LIST  (read-only, no approval needed)
# ==============================================================================

def cmd_list(user: dict) -> None:
    _banner(f"MEMBERSHIP STATUS — {user.get('displayName')}")
    for r in _all_resources():
        try:
            member = is_member(r, user["id"], user["mail"] or user["userPrincipalName"])
            status = "MEMBER" if member else "not a member"
        except requests.HTTPError as exc:
            status = f"[error checking: {_safe_error(exc)}]"
        print(f"  [{r['type'].upper():<10}] {r['label']:<28} → {status}")


# ==============================================================================
# AUTOMATIC CHECK — every company group is checked, missing ones are queued
# ==============================================================================

def cmd_auto_check(user: dict) -> tuple[list[dict], list[dict]]:
    """
    Checks the employee against EVERY configured company group/DL/Team/Site.
    Returns (already_member, missing) so the rest of the flow can act on it.
    Nothing is written to Graph here — this is read-only.
    """
    _banner(f"CHECKING ALL COMPANY GROUPS — {user.get('displayName')}")
    already_member, missing = [], []
    for r in _all_resources():
        try:
            member = is_member(r, user["id"], user["mail"] or user["userPrincipalName"])
        except requests.HTTPError as exc:
            print(f"  [{r['type'].upper():<10}] {r['label']:<28} → [error checking: {_safe_error(exc)}]")
            continue
        status = "MEMBER" if member else "MISSING"
        print(f"  [{r['type'].upper():<10}] {r['label']:<28} → {status}")
        (already_member if member else missing).append(r)
    return already_member, missing


def cmd_add_missing(user: dict, missing: list[dict]) -> None:
    """Queues an ADD approval request for every group the employee is missing."""
    if not missing:
        print("\n  Already a member of every configured group — nothing to add.")
        return
    print(f"\n  {len(missing)} group(s) missing:")
    for r in missing:
        print(f"    - [{r['type'].upper()}] {r['label']}")
    answer = input("\n  Request owner approval to add this employee to ALL of the above? (y/n): ").strip().lower()
    if answer != "y":
        print("  Skipped — no requests created.")
        return
    for r in missing:
        queue_request("ADD", r, user)
    print("\n  All ADD requests are PENDING owner approval. Nothing has been changed yet.")


# ==============================================================================
# REMOVE  (membership only, group/team never deleted) — offered as a follow-up
# ==============================================================================

def cmd_remove(user: dict, current_member: list[dict]) -> None:
    if not current_member:
        return
    answer = input("\n  Does this employee need to be REMOVED from any group above? (y/n): ").strip().lower()
    if answer != "y":
        return
    _banner("REMOVE — select resource(s) to request membership removal for")
    print("  (This removes MEMBERSHIP only. The Group/Team/DL/Site itself is never deleted.)")
    resources = _resource_menu(current_member)
    if not resources:
        print("  No valid selection made.")
        return
    for r in resources:
        queue_request("REMOVE", r, user)
    print("\n  All REMOVE requests are PENDING owner approval. Nothing has been changed yet.")


# ==============================================================================
# MODE 4 — MOVE TEAM  (remove old team + add new team, linked, both pending)
# ==============================================================================

def cmd_move_team(user: dict) -> None:
    _banner("MOVE TEAM — remove from old team, add to new team")
    teams = [r for r in _all_resources() if r["type"] == "team"]

    print("\n  Which team is the employee LEAVING?")
    old_team = _pick_single(teams)
    if not old_team:
        print("  Cancelled — no valid team selected.")
        return

    remaining = [t for t in teams if t["id"] != old_team["id"]]
    print("\n  Which team is the employee JOINING?")
    new_team = _pick_single(remaining)
    if not new_team:
        print("  Cancelled — no valid team selected.")
        return

    reason = f"Team move: {old_team['label']} → {new_team['label']}"

    if is_member(old_team, user["id"], user["mail"] or user["userPrincipalName"]):
        queue_request("REMOVE", old_team, user, reason=reason)
    else:
        print(f"  [SKIP] Not currently a member of {old_team['label']} — no removal request needed.")

    if not is_member(new_team, user["id"], user["mail"] or user["userPrincipalName"]):
        queue_request("ADD", new_team, user, reason=reason)
    else:
        print(f"  [SKIP] Already a member of {new_team['label']} — no add request needed.")

    print("\n  Both requests are PENDING owner approval. Nothing has been changed yet.")


def _pick_single(resources: list[dict]) -> dict | None:
    for i, r in enumerate(resources, start=1):
        print(f"    {i:>2}) {r['label']}")
    raw = input("    Enter number: ").strip()
    if raw.isdigit() and 1 <= int(raw) <= len(resources):
        return resources[int(raw) - 1]
    return None


# ==============================================================================
# MODE 5 — APPROVE  (owner only — this is the only place Graph writes happen)
# ==============================================================================

def cmd_approve() -> None:
    pending = [r for r in load_pending() if r["status"] == "PENDING"]
    if not pending:
        print("\n  No pending requests. Nothing to approve.")
        return

    _banner(f"APPROVAL QUEUE — {len(pending)} pending request(s)")
    all_requests = load_pending()

    for req in pending:
        print(f"\n  Request {req['id']}  ({req['created_at']})")
        print(f"    Employee : {req['employee_name']} ({req['employee_email']})")
        print(f"    Action   : {req['action']}")
        print(f"    Resource : {req['resource_label']}  [{req['resource_type']}]")
        if req.get("reason"):
            print(f"    Reason   : {req['reason']}")
        if req["resource_type"] in ("group", "dl", "team"):
            print(f"    Note     : membership only — the {req['resource_type']} will NOT be deleted")

        answer = input("    Approve this? (y/n/skip): ").strip().lower()
        if answer == "y":
            try:
                execute_request(req)
                req["status"] = "EXECUTED"
                verb = "Added" if req["action"] == "ADD" else "Removed"
                print(f"    [DONE ✓] {verb} successfully.")
            except requests.HTTPError as exc:
                req["status"] = "FAILED"
                print(f"    [FAILED] HTTP {exc.response.status_code} — {_safe_error(exc)}")
            except Exception as exc:
                req["status"] = "FAILED"
                print(f"    [FAILED] {exc}")
        elif answer == "n":
            req["status"] = "DENIED"
            print("    [DENIED] Request closed, no change made.")
        else:
            print("    [SKIPPED] Left pending for next time.")
            continue

        # write back status for this request id
        for r in all_requests:
            if r["id"] == req["id"]:
                r["status"] = req["status"]
        save_pending(all_requests)

    print("\n  Approval session complete.")


# ==============================================================================
# ENTRY POINT
# ==============================================================================

def _get_email_and_user() -> dict | None:
    email = input("\nEnter Employee Email: ").strip()
    if not email or "@" not in email:
        print("[ERROR] Invalid email address.")
        return None
    print(f"\n[•] Looking up user: {email} ...")
    user = resolve_user(email)
    if not user:
        print(f"\n[ERROR] '{email}' was not found in Azure Active Directory.")
        return None
    print(f"  [FOUND] {user.get('displayName')}  ({user['id']})")
    return user


def main() -> None:
    print("\n" + "=" * 62)
    print("   Microsoft 365  —  Membership Manager (with Owner Approval)")
    print("=" * 62)

    _require_env("MS_TENANT_ID")
    _require_env("MS_CLIENT_ID")
    _require_env("MS_CLIENT_SECRET")

    print("\n[•] Authenticating with Microsoft Graph ...")
    try:
        get_access_token()
        print("  [SUCCESS] Authenticated successfully")
    except Exception as exc:
        print(f"\n[ERROR] Authentication failed: {exc}")
        sys.exit(1)

    owner_mode = input("\nAre you the owner approving pending requests? (y/n): ").strip().lower()
    if owner_mode == "y":
        cmd_approve()
        return

    # ── Normal flow: just give the email, everything else is automatic ──────
    user = _get_email_and_user()
    if not user:
        sys.exit(1)

    already_member, missing = cmd_auto_check(user)
    cmd_add_missing(user, missing)
    cmd_remove(user, already_member)

    move_answer = input("\n  Is this employee moving from one team to another? (y/n): ").strip().lower()
    if move_answer == "y":
        cmd_move_team(user)

    print("\n  Done. Anything queued above is PENDING until the owner runs this")
    print("  script again and answers 'y' to the approval question at the start.")


if __name__ == "__main__":
    main()