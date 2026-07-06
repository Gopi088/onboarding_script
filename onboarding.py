"""
onboarding.py
================
Microsoft 365 Employee Onboarding Automation
Uses Microsoft Graph API to check and add a user to all configured resources.

Authentication (client-credentials OAuth2 via MSAL):
  Set these environment variables before running:
    MS_TENANT_ID      – Your Azure AD Tenant ID
    MS_CLIENT_ID      – App (client) ID of your Azure AD App Registration
    MS_CLIENT_SECRET  – Client Secret for that app

  Required Microsoft Graph Application Permissions (granted in Azure AD):
    User.Read.All
    Group.ReadWrite.All
    GroupMember.ReadWrite.All
    TeamMember.ReadWrite.All
    Sites.FullControl.All
    Directory.ReadWrite.All

Usage:
  pip install msal requests python-dotenv
  python onboarding.py
"""

import os
import sys
import time
import requests
from dotenv import load_dotenv

# Load .env file if present (safe to call even if file doesn't exist)
load_dotenv()

try:
    import msal
except ImportError:
    print("[ERROR] msal is not installed. Run: pip install msal")
    sys.exit(1)


# ==============================================================================
# ██████╗ ███████╗███████╗ ██████╗ ██╗   ██╗██████╗  ██████╗███████╗███████╗
# ██╔══██╗██╔════╝██╔════╝██╔═══██╗██║   ██║██╔══██╗██╔════╝██╔════╝██╔════╝
# ██████╔╝█████╗  ███████╗██║   ██║██║   ██║██████╔╝██║     █████╗  ███████╗
# ██╔══██╗██╔══╝  ╚════██║██║   ██║██║   ██║██╔══██╗██║     ██╔══╝  ╚════██║
# ██║  ██║███████╗███████║╚██████╔╝╚██████╔╝██║  ██║╚██████╗███████╗███████║
# ╚═╝  ╚═╝╚══════╝╚══════╝ ╚═════╝  ╚═════╝ ╚═╝  ╚═╝ ╚═════╝╚══════╝╚══════╝
#
# !! EDIT ONLY THIS SECTION TO CONFIGURE YOUR ENVIRONMENT !!
# Replace every placeholder value with your real Azure AD Object IDs.
#
# How to find IDs:
#   Groups / DLs   → Azure AD Portal → Groups → click group → copy Object ID
#   Teams          → Azure AD Portal → Groups → filter by "Teams" → Object ID
#   SharePoint     → Run: GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{name}
# ==============================================================================

# ── Microsoft 365 Groups (and Security Groups) ────────────────────────────────
# These are standard Azure AD groups. The script will add the user as a member.
M365_GROUPS: dict[str, str] = {
    # "Display Name (for logs)": "Azure AD Object ID"
    "Everyone Group":            "00000000-0000-0000-0000-000000000001",
    "Microsoft 365 Main Group":  "00000000-0000-0000-0000-000000000002",
    "Security Group - General":  "00000000-0000-0000-0000-000000000003",
}

# ── Distribution Lists ────────────────────────────────────────────────────────
# Distribution lists in Exchange Online also appear as Azure AD groups.
DISTRIBUTION_LISTS: dict[str, str] = {
    # "Display Name (for logs)": "Azure AD Object ID"
    "All Staff DL":              "00000000-0000-0000-0000-000000000010",
    "Company Announcements DL":  "00000000-0000-0000-0000-000000000011",
}

# ── Microsoft Teams ───────────────────────────────────────────────────────────
# Each Teams entry needs the underlying Microsoft 365 Group Object ID
# (Teams are backed by M365 Groups — same ID, different API endpoint).
MS_TEAMS: dict[str, str] = {
    # "Display Name (for logs)": "Azure AD Object ID of the backing M365 Group"
    "Alpha Team":                "00000000-0000-0000-0000-000000000020",
    "Beta Team":                 "00000000-0000-0000-0000-000000000021",
    "Gamma Team":                "00000000-0000-0000-0000-000000000022",
    "TA Team":                   "00000000-0000-0000-0000-000000000023",
    "TL Team":                   "00000000-0000-0000-0000-000000000024",
    "Software Engineer Team":    "00000000-0000-0000-0000-000000000025",
}

# ── SharePoint Sites ──────────────────────────────────────────────────────────
# Format: "Label": ("Graph Site ID", "Permission role")
# Site ID format: "{hostname},{site-collection-id},{web-id}"
# Fetch it via: GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{name}
# Roles: "read" | "write" | "fullControl"
SHAREPOINT_SITES: dict[str, tuple[str, str]] = {
    # "Display Name (for logs)": ("Graph Site ID", "role")
    "Home SharePoint":        ("contoso.sharepoint.com,aaa-bbb,ccc-ddd", "read"),
    "Recruitment SharePoint": ("contoso.sharepoint.com,eee-fff,ggg-hhh", "read"),
}

# ── SharePoint hostname (used when checking existing permissions) ─────────────
# Example: "contoso.sharepoint.com"
SHAREPOINT_HOSTNAME: str = os.environ.get("SHAREPOINT_HOSTNAME", "contoso.sharepoint.com")

# ==============================================================================
# END OF CONFIGURATION
# ==============================================================================


# ── Graph API base URL ────────────────────────────────────────────────────────
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# ── In-process token cache ────────────────────────────────────────────────────
_token_cache: dict = {}

# ── Run-level result counters ─────────────────────────────────────────────────
STATS = {"completed": 0, "skipped": 0, "failed": 0}


# ==============================================================================
# AUTHENTICATION
# ==============================================================================

def get_access_token() -> str:
    """
    Acquire (or return cached) a Microsoft Graph access token
    using the client-credentials OAuth2 flow via MSAL.
    """
    now = time.time()
    if _token_cache.get("expires_at", 0) > now + 60:
        return _token_cache["token"]

    tenant_id     = _require_env("MS_TENANT_ID")
    client_id     = _require_env("MS_CLIENT_ID")
    client_secret = _require_env("MS_CLIENT_SECRET")

    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=authority,
        client_credential=client_secret,
    )

    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in result:
        err = result.get("error_description") or result.get("error") or str(result)
        raise RuntimeError(f"Failed to acquire access token: {err}")

    _token_cache["token"]      = result["access_token"]
    _token_cache["expires_at"] = now + result.get("expires_in", 3600)
    return _token_cache["token"]


def _require_env(name: str) -> str:
    """Return an environment variable or exit with a clear message."""
    value = os.environ.get(name, "").strip()
    if not value:
        print(f"\n[ERROR] Environment variable '{name}' is not set.")
        print("        Please set it before running this script.")
        sys.exit(1)
    return value


def _headers() -> dict:
    """Return Authorization headers for Graph API calls."""
    return {
        "Authorization": f"Bearer {get_access_token()}",
        "Content-Type":  "application/json",
    }


# ==============================================================================
# GRAPH API — GENERIC HELPERS
# ==============================================================================

def graph_get(path: str, params: dict = None) -> dict:
    """Perform a GET request against Microsoft Graph. Returns parsed JSON."""
    url = f"{GRAPH_BASE}{path}"
    resp = requests.get(url, headers=_headers(), params=params, timeout=30)
    resp.raise_for_status()
    return resp.json()


def graph_post(path: str, payload: dict) -> dict:
    """Perform a POST request against Microsoft Graph. Returns parsed JSON."""
    url = f"{GRAPH_BASE}{path}"
    resp = requests.post(url, headers=_headers(), json=payload, timeout=30)
    resp.raise_for_status()
    return resp.json() if resp.content else {}


# ==============================================================================
# USER RESOLUTION
# ==============================================================================

def resolve_user(email: str) -> dict | None:
    """
    Look up a user by email using Microsoft Graph.
    Returns the user object (with 'id', 'displayName', etc.) or None if not found.
    """
    try:
        data = graph_get(f"/users/{email}", params={"$select": "id,displayName,mail,userPrincipalName"})
        return data
    except requests.HTTPError as exc:
        if exc.response.status_code == 404:
            return None
        raise


# ==============================================================================
# MEMBERSHIP CHECK HELPERS
# ==============================================================================

def is_group_member(user_id: str, group_id: str) -> bool:
    """
    Check if a user is a direct member of an Azure AD group.
    Uses: GET /groups/{group-id}/members/{user-id}
    Returns True if member, False otherwise.
    """
    try:
        graph_get(f"/groups/{group_id}/members/{user_id}")
        return True
    except requests.HTTPError as exc:
        if exc.response.status_code in (404, 400):
            return False
        raise


def is_team_member(user_id: str, team_id: str) -> bool:
    """
    Check if a user is a member of a Microsoft Team.
    Lists all members and checks for a match — Graph does not supportF
    a single-user lookup on /teams/{id}/members directly.
    """
    try:
        data = graph_get(f"/teams/{team_id}/members")
        members = data.get("value", [])
        return any(m.get("userId") == user_id for m in members)
    except requests.HTTPError as exc:
        if exc.response.status_code in (404, 400):
            return False
        raise


def has_sharepoint_access(user_email: str, site_id: str) -> bool:
    """
    Check if a user already has a permission entry on a SharePoint site.
    Lists site permissions and looks for the user's email.
    """
    try:
        data = graph_get(f"/sites/{site_id}/permissions")
        perms = data.get("value", [])
        for perm in perms:
            granted = perm.get("grantedToIdentities") or []
            for identity in granted:
                if identity.get("user", {}).get("email", "").lower() == user_email.lower():
                    return True
        return False
    except requests.HTTPError:
        # If we can't read permissions, assume not present and try to add.
        return False


# ==============================================================================
# ADD MEMBER HELPERS
# ==============================================================================

def add_to_group(user_id: str, group_id: str) -> None:
    """
    Add a user as a member of an Azure AD group (M365 Group, Security Group,
    or Distribution List).
    Uses: POST /groups/{group-id}/members/$ref
    """
    payload = {
        "@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{user_id}"
    }
    graph_post(f"/groups/{group_id}/members/$ref", payload)


def add_to_team(user_id: str, team_id: str) -> None:
    """
    Add a user as a member of a Microsoft Team.
    Uses: POST /teams/{team-id}/members
    """
    payload = {
        "@odata.type": "#microsoft.graph.aadUserConversationMember",
        "roles": [],
        "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{user_id}')",
    }
    graph_post(f"/teams/{team_id}/members", payload)


def grant_sharepoint(user_email: str, site_id: str, role: str) -> None:
    """
    Grant a user access to a SharePoint site.
    Uses: POST /sites/{site-id}/permissions
    role: "read" | "write" | "fullControl"
    """
    payload = {
        "roles": [role],
        "grantedToIdentities": [
            {"user": {"email": user_email}}
        ],
    }
    graph_post(f"/sites/{site_id}/permissions", payload)


# ==============================================================================
# PRINT HELPERS
# ==============================================================================

def _banner(text: str) -> None:
    print(f"\n{'─' * 55}")
    print(f"  {text}")
    print(f"{'─' * 55}")


def _log(label: str, msg: str) -> None:
    print(f"  {label:<14} {msg}")


# ==============================================================================
# RESOURCE PROCESSORS
# ==============================================================================

def process_group(user_id: str, label: str, group_id: str) -> None:
    """Check and optionally add a user to an Azure AD group or Distribution List."""
    print(f"\n[•] Checking {label} ...")
    try:
        if is_group_member(user_id, group_id):
            _log("[SKIPPED]", "Already a member")
            STATS["skipped"] += 1
        else:
            _log("[INFO]", "Not a member — adding ...")
            add_to_group(user_id, group_id)
            _log("[SUCCESS]", "Added successfully")
            STATS["completed"] += 1
    except requests.HTTPError as exc:
        _log("[FAILED]", f"{exc.response.status_code} — {_safe_error(exc)}")
        STATS["failed"] += 1
    except Exception as exc:
        _log("[FAILED]", str(exc))
        STATS["failed"] += 1


def process_team(user_id: str, label: str, team_id: str) -> None:
    """Check and optionally add a user to a Microsoft Team."""
    print(f"\n[•] Checking {label} ...")
    try:
        if is_team_member(user_id, team_id):
            _log("[SKIPPED]", "Already a member")
            STATS["skipped"] += 1
        else:
            _log("[INFO]", "Not a member — adding ...")
            add_to_team(user_id, team_id)
            _log("[SUCCESS]", "Added successfully")
            STATS["completed"] += 1
    except requests.HTTPError as exc:
        _log("[FAILED]", f"{exc.response.status_code} — {_safe_error(exc)}")
        STATS["failed"] += 1
    except Exception as exc:
        _log("[FAILED]", str(exc))
        STATS["failed"] += 1


def process_sharepoint(user_email: str, label: str, site_id: str, role: str) -> None:
    """Check and optionally grant a user access to a SharePoint site."""
    print(f"\n[•] Checking {label} ...")
    try:
        if has_sharepoint_access(user_email, site_id):
            _log("[SKIPPED]", "Already has access")
            STATS["skipped"] += 1
        else:
            _log("[INFO]", f"No access found — granting '{role}' permission ...")
            grant_sharepoint(user_email, site_id, role)
            _log("[SUCCESS]", "Access granted successfully")
            STATS["completed"] += 1
    except requests.HTTPError as exc:
        _log("[FAILED]", f"{exc.response.status_code} — {_safe_error(exc)}")
        STATS["failed"] += 1
    except Exception as exc:
        _log("[FAILED]", str(exc))
        STATS["failed"] += 1


def _safe_error(exc: requests.HTTPError) -> str:
    """Extract a readable message from an HTTPError response body."""
    try:
        body = exc.response.json()
        return body.get("error", {}).get("message", exc.response.text[:120])
    except Exception:
        return exc.response.text[:120]


# ==============================================================================
# MAIN ORCHESTRATOR
# ==============================================================================

def run_onboarding(email: str) -> None:
    """
    Full onboarding workflow for a single employee email:
      1. Resolve user → get Object ID
      2. Loop through every configured resource
      3. Check membership / access
      4. Add only if missing
      5. Print final summary
    """
    # ── Step 1: Resolve user ──────────────────────────────────────────────────
    print(f"\n[•] Looking up user: {email} ...")
    user = resolve_user(email)

    if not user:
        print(f"\n[ERROR] User '{email}' was not found in Azure Active Directory.")
        print("        Please verify the email address and try again.")
        sys.exit(1)

    user_id      = user["id"]
    display_name = user.get("displayName") or user.get("userPrincipalName") or email

    print(f"  [FOUND]  {display_name}")
    print(f"  Object ID: {user_id}")

    # ── Step 2: Process Microsoft 365 Groups & Security Groups ────────────────
    _banner("MICROSOFT 365 GROUPS & SECURITY GROUPS")
    for label, group_id in M365_GROUPS.items():
        process_group(user_id, label, group_id)

    # ── Step 3: Process Distribution Lists ───────────────────────────────────
    _banner("DISTRIBUTION LISTS")
    for label, group_id in DISTRIBUTION_LISTS.items():
        process_group(user_id, label, group_id)

    # ── Step 4: Process Microsoft Teams ──────────────────────────────────────
    _banner("MICROSOFT TEAMS")
    for label, team_id in MS_TEAMS.items():
        process_team(user_id, label, team_id)

    # ── Step 5: Process SharePoint Sites ─────────────────────────────────────
    _banner("SHAREPOINT SITES")
    for label, (site_id, role) in SHAREPOINT_SITES.items():
        process_sharepoint(email, label, site_id, role)

    # ── Step 6: Final Summary ─────────────────────────────────────────────────
    total = STATS["completed"] + STATS["skipped"] + STATS["failed"]
    print(f"\n{'=' * 55}")
    print(f"  ONBOARDING COMPLETE — {display_name}")
    print(f"{'=' * 55}")
    print(f"  Total resources checked : {total}")
    print(f"  ✓  Completed (added)    : {STATS['completed']}")
    print(f"  ⊘  Skipped (existing)   : {STATS['skipped']}")
    print(f"  ✗  Failed               : {STATS['failed']}")
    print(f"{'=' * 55}\n")


# ==============================================================================
# ENTRY POINT
# ==============================================================================

def main() -> None:
    print("\n" + "=" * 55)
    print("   Microsoft 365 Onboarding Automation")
    print("=" * 55)

    email = input("\nEnter Employee Email: ").strip().lower()

    if not email or "@" not in email:
        print("[ERROR] Invalid email address entered.")
        sys.exit(1)

    # Verify credentials are present before doing any work
    _require_env("MS_TENANT_ID")
    _require_env("MS_CLIENT_ID")
    _require_env("MS_CLIENT_SECRET")

    # Acquire token early so auth errors surface immediately
    print("\n[•] Authenticating with Microsoft Graph ...")
    try:
        get_access_token()
        print("  [SUCCESS] Authentication successful")
    except Exception as exc:
        print(f"\n[ERROR] Authentication failed: {exc}")
        sys.exit(1)

    run_onboarding(email)


if __name__ == "__main__":
    main()