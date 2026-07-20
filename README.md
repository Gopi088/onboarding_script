# team-member

Microsoft 365 group management CLI tool — create groups, list/add/remove an employee's group membership.

## Setup

```bash
pip install msal requests python-dotenv
```

Azure AD App Registration needs these **Application permissions** with admin consent granted:

- `User.Read.All`
- `Group.ReadWrite.All`
- `GroupMember.ReadWrite.All`

Create a `.env` file in the directory you run the command from:

```
MS_TENANT_ID=your-tenant-id
MS_CLIENT_ID=your-client-id
MS_CLIENT_SECRET=your-client-secret
```

## Install

```bash
chmod +x team-member
sudo cp team-member /usr/local/bin/team-member
```

## Usage

```
team-member create-group --name "Everyone"
team-member create-group --name "Hiring" --type unified --description "Hiring team DL"
team-member list --email user@company.com
team-member add --email user@company.com --group "Everyone"
team-member remove --email user@company.com --group "Everyone"
```

### Options

| Flag | Description |
|---|---|
| `--verbose` | Print logs to terminal (off by default; full detail always goes to `team-member.log`) |
| `--log-level LEVEL` | One of `DEBUG`, `INFO`, `WARNING`, `ERROR` (default: `INFO`; only matters with `--verbose`) |

### Commands

| Command | Description |
|---|---|
| `create-group --name NAME [--type security\|unified] [--description DESC] [--force]` | Create a security or Microsoft 365 group. `--force` allows duplicate names. |
| `list --email EMAIL` | List all groups the user is a member of. |
| `add --email EMAIL --group GROUP` | Add user to a group (membership only). |
| `remove --email EMAIL --group GROUP` | Remove user from a group (membership only — the group itself is never deleted). |

## How it works

Three-layer design:

- **Layer 1 — GraphClient**: Raw HTTP/auth calls to Microsoft Graph. Knows nothing about groups or employees.
- **Layer 2 — GroupService**: Business logic (create group, list membership, add/remove member). Built only on `GraphClient`.
- **Layer 3 — CLI**: `argparse` commands + logging. The only part that touches `sys.argv` or `print()`.

## Notes

- Full operation logs are always written to `team-member.log` in the current directory.
- Microsoft Graph has eventual consistency. After `add` or `remove`, the tool polls until the change is confirmed before returning, so a subsequent `list` shows correct data.
