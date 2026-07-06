# Microsoft 365 Employee Onboarding Automation

**Script:** `onboarding.py`  
**Purpose:** Automatically checks and provisions a new employee's Microsoft 365 access — Groups, Distribution Lists, Teams, and SharePoint — in a single run.

---

## What It Does

When you run this script with a new employee's email address, it:

1. **Looks up** the user in Azure Active Directory
2. **Checks** whether they already have access to each configured resource
3. **Adds** them only where they are missing (skips if already a member)
4. **Reports** a final summary of what was added, skipped, or failed

No access is ever removed. The script is fully **idempotent** — running it twice on the same employee produces the same result safely.

---

## Resources Provisioned

| Category | What Gets Configured |
|---|---|
| **M365 Groups / Security Groups** | Adds user as a group member |
| **Distribution Lists** | Adds user to all company email lists |
| **Microsoft Teams** | Adds user as a member of each configured Team |
| **SharePoint Sites** | Grants read/write access to configured sites |

---

## Prerequisites

### 1. Python & Packages
```bash
pip install msal requests python-dotenv
```
Requires **Python 3.10+** (uses modern type hints).

### 2. Azure AD App Registration
Your Azure AD App needs these **Application Permissions** (not delegated):

| Permission | Used For |
|---|---|
| `User.Read.All` | Look up user by email |
| `Group.ReadWrite.All` | Check & add group members |
| `GroupMember.ReadWrite.All` | Modify group membership |
| `TeamMember.ReadWrite.All` | Add members to Teams |
| `Sites.FullControl.All` | Grant SharePoint permissions |
| `Directory.ReadWrite.All` | General directory operations |

> After adding permissions in Azure AD, click **"Grant admin consent"**.

---

## Setup

### Step 1 — Set Environment Variables

Create a `.env` file in the same folder as the script (or export them in your terminal):

```env
MS_TENANT_ID=your-azure-tenant-id
MS_CLIENT_ID=your-app-client-id
MS_CLIENT_SECRET=your-app-client-secret
SHAREPOINT_HOSTNAME=yourcompany.sharepoint.com
```

> **Where to find these values:**
> - `MS_TENANT_ID` → Azure AD Portal → Overview → Tenant ID
> - `MS_CLIENT_ID` → Azure AD → App Registrations → your app → Application (client) ID
> - `MS_CLIENT_SECRET` → Azure AD → App Registrations → your app → Certificates & secrets

---

### Step 2 — Configure Your Resources

Open `onboarding.py` and edit **only the configuration section** at the top of the file (clearly marked with a banner comment).

#### M365 Groups & Security Groups
```python
M365_GROUPS: dict[str, str] = {
    "Everyone Group":  "paste-azure-object-id-here",
    "Security Group":  "paste-azure-object-id-here",
}
```

#### Distribution Lists
```python
DISTRIBUTION_LISTS: dict[str, str] = {
    "All Staff DL":  "paste-azure-object-id-here",
}
```

#### Microsoft Teams
```python
MS_TEAMS: dict[str, str] = {
    "Engineering Team":  "paste-azure-object-id-here",
}
```
> The ID here is the **backing M365 Group Object ID** of the Team, not a separate ID.

#### SharePoint Sites
```python
SHAREPOINT_SITES: dict[str, tuple[str, str]] = {
    "Intranet":  ("yourcompany.sharepoint.com,site-collection-id,web-id", "read"),
}
```
> **Role options:** `"read"` | `"write"` | `"fullControl"`  
> **How to get the Site ID:** Call `GET https://graph.microsoft.com/v1.0/sites/yourcompany.sharepoint.com:/sites/sitename`

---

### Step 3 — How to Find Object IDs

| Resource | Where to Find the ID |
|---|---|
| Groups / DLs | Azure AD Portal → Groups → click the group → **Object ID** |
| Teams | Azure AD Portal → Groups → filter by "Teams" → **Object ID** |
| SharePoint | Graph Explorer: `GET /v1.0/sites/{hostname}:/sites/{name}` |

---

## Running the Script

```bash
python onboarding.py
```

You will be prompted to enter the employee's email:
```
=======================================================
   Microsoft 365 Onboarding Automation
=======================================================

Enter Employee Email: john.doe@yourcompany.com
```

### Example Output

```
[•] Looking up user: john.doe@yourcompany.com ...
  [FOUND]   John Doe
  Object ID: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

───────────────────────────────────────────────────────
  MICROSOFT 365 GROUPS & SECURITY GROUPS
───────────────────────────────────────────────────────

[•] Checking Everyone Group ...
  [SKIPPED]      Already a member

[•] Checking Security Group - General ...
  [INFO]         Not a member — adding ...
  [SUCCESS]      Added successfully

───────────────────────────────────────────────────────
  MICROSOFT TEAMS
───────────────────────────────────────────────────────

[•] Checking Engineering Team ...
  [INFO]         Not a member — adding ...
  [SUCCESS]      Added successfully

=======================================================
  ONBOARDING COMPLETE — John Doe
=======================================================
  Total resources checked : 11
  ✓  Completed (added)    : 2
  ⊘  Skipped (existing)   : 8
  ✗  Failed               : 1
=======================================================
```

---

## Status Meanings

| Status | Meaning |
|---|---|
| `[SUCCESS]` | User was not a member — successfully added |
| `[SKIPPED]` | User was already a member — no action taken |
| `[FAILED]` | An API error occurred — check permissions or the ID |

---

## Test Results

The script has been validated with **25 unit tests** covering all critical paths:

| # | Test Case | Result |
|---|---|---|
| TC-01 | Token caching (no redundant auth calls) | ✅ PASS |
| TC-02 | Expired token is automatically re-acquired | ✅ PASS |
| TC-03 | User lookup returns correct object when found | ✅ PASS |
| TC-04 | User lookup returns None for unknown email (404) | ✅ PASS |
| TC-05 | Group membership check returns True for existing member | ✅ PASS |
| TC-06 | Group membership check returns False when not a member | ✅ PASS |
| TC-07 | Team membership check returns True for existing member | ✅ PASS |
| TC-08 | Team membership check returns False when absent | ✅ PASS |
| TC-09 | SharePoint access detected when permission entry found | ✅ PASS |
| TC-10 | SharePoint email check is case-insensitive | ✅ PASS |
| TC-11 | SharePoint API error handled safely (returns False) | ✅ PASS |
| TC-12 | Group: skips add if user is already a member | ✅ PASS |
| TC-13 | Group: adds user and increments completed count | ✅ PASS |
| TC-14 | Group: records FAILED stat on API error | ✅ PASS |
| TC-15 | Team: skips add if user is already a member | ✅ PASS |
| TC-16 | Team: adds user and increments completed count | ✅ PASS |
| TC-17 | SharePoint: skips grant if user already has access | ✅ PASS |
| TC-18 | SharePoint: grants access for new user | ✅ PASS |
| TC-19 | add_to_group sends correct OData payload | ✅ PASS |
| TC-20 | add_to_team sends correct member payload | ✅ PASS |
| TC-21 | grant_sharepoint sends correct role and identity | ✅ PASS |
| TC-22 | Exits with error on email missing `@` | ✅ PASS |
| TC-23 | Exits with error on empty email input | ✅ PASS |
| TC-24 | Stats accumulate correctly across multiple resources | ✅ PASS |
| TC-25 | API error messages extracted and displayed clearly | ✅ PASS |

**All 25/25 tests passed.**

---

## Common Errors & Fixes

| Error Message | Cause | Fix |
|---|---|---|
| `Environment variable 'MS_TENANT_ID' is not set` | Missing `.env` or env vars | Set all 3 required env vars |
| `Failed to acquire access token` | Wrong client ID/secret | Double-check credentials in Azure AD |
| `[FAILED] 403` on a resource | App missing API permission | Add the missing permission + grant admin consent |
| `[FAILED] 404` on a resource | Wrong Object ID configured | Verify ID in Azure AD Portal |
| `User not found in Azure Active Directory` | Email doesn't exist in tenant | Confirm user account has been created first |

---

## Security Notes

- The script uses **client credentials flow** (app-level auth, no user login required).
- Client secrets should be stored in environment variables or a secrets manager — **never hardcoded**.
- The app registration should follow the principle of least privilege; only grant the permissions listed above.
- Token is cached in-memory for the duration of the run only.

---

## File Structure

```
onboarding.py     ← Main script (edit the config section at the top)
.env                 ← Your credentials (never commit to source control)
README.md            ← This file
```

---

*For questions or to add new resources (new Teams, Groups, SharePoint sites), edit the configuration dictionaries at the top of `onboarding.py` following the existing format.*