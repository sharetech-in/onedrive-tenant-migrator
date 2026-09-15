# OneDrive tenant-to-tenant folder migration

This utility copies a folder tree from one user's OneDrive to another user's OneDrive through Microsoft Graph. It uses two app registrations, one in each tenant, and never stores secrets in the repository.

## What It Does

- Authenticates to each tenant with Microsoft Entra app-only client credentials.
- Reads the source user's OneDrive folder tree.
- Creates missing destination folders.
- Copies files recursively, including large files through Graph upload sessions.
- Replaces destination files with the same name.
- Retries transient Graph responses such as throttling and temporary server errors.
- Saves completed source item IDs in `migration-state.json` so an interrupted migration can be resumed.

## Requirements

- Python 3.9 or newer
- Two app registrations, one in the source tenant and one in the destination tenant
- Admin consent for the required Microsoft Graph application permissions
- Access to the source and destination users' OneDrive accounts

## Microsoft Graph Permissions

Configure these permissions under **Microsoft Graph > Application permissions**:

| App registration | Permission | Purpose |
| --- | --- | --- |
| Source tenant | `Files.Read.All` | List folders and download source files |
| Destination tenant | `Files.ReadWrite.All` | Create folders and upload or replace files |

Grant admin consent for both applications. Delegated permissions and `User.Read` are not required because the script uses the client-credentials flow.

## Configuration

Copy the safe template and edit the local copy:

```bash
cp .env.example .env
```

Set these values in `.env`:

```dotenv
SOURCE_TENANT_ID=<source-tenant-id>
SOURCE_CLIENT_ID=<source-app-client-id>
SOURCE_CLIENT_SECRET=<source-app-client-secret>
SOURCE_FOLDER_PATH=Images/Documents

DEST_TENANT_ID=<destination-tenant-id>
DEST_CLIENT_ID=<destination-app-client-id>
DEST_CLIENT_SECRET=<destination-app-client-secret>
DESTINATION_FOLDER_PATH=Migration/Documents
```

The folder paths are relative to each user's OneDrive root. Do not include the `My Files` label shown in the OneDrive web interface. For example:

```dotenv
SOURCE_FOLDER_PATH=Malvan Trip
DESTINATION_FOLDER_PATH=Migration/Malvan Trip
```

Use an empty path to select the OneDrive root:

```dotenv
SOURCE_FOLDER_PATH=
DESTINATION_FOLDER_PATH=
```

The destination path is created automatically during a real migration if it does not already exist. A dry run does not create destination folders.

## Installation

Create and activate a virtual environment, then install dependencies:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Use the same `.venv/bin/python` interpreter for subsequent commands if the environment is not activated.

## Dry Run

A dry run authenticates to Graph, reads the source tree, and logs the files that would be copied. It does not create folders, upload files, or save migration state.

```bash
.venv/bin/python migrate_onedrive.py \
  --source-user source.user@example.com \
  --destination-user destination.user@example.com \
  --dry-run
```

## Run the Migration

After reviewing the dry-run output, remove `--dry-run`:

```bash
.venv/bin/python migrate_onedrive.py \
  --source-user source.user@example.com \
  --destination-user destination.user@example.com
```

The source and destination folder paths are read from `.env`. They can be overridden for a single run:

```bash
.venv/bin/python migrate_onedrive.py \
  --source-user source.user@example.com \
  --destination-user destination.user@example.com \
  --source-folder "Images/Documents" \
  --destination-folder "Migration/Documents"
```

## Command-Line Options

| Option | Required | Description |
| --- | --- | --- |
| `--source-user` | Yes | Source user's UPN or object ID |
| `--destination-user` | Yes | Destination user's UPN or object ID |
| `--source-folder` | No | Overrides `SOURCE_FOLDER_PATH` |
| `--destination-folder` | No | Overrides `DESTINATION_FOLDER_PATH` |
| `--state` | No | State file path; defaults to `migration-state.json` |
| `--dry-run` | No | List work without changing the destination |
| `--verbose` | No | Enable debug logging |

## Resume Behavior

During a real migration, each completed source item ID is saved to `migration-state.json`. If the process is interrupted, rerun the same command and completed items are skipped.

Use a separate state file for separate migrations:

```bash
.venv/bin/python migrate_onedrive.py \
  --source-user source.user@example.com \
  --destination-user destination.user@example.com \
  --state migration-malvan.json
```

State files are local migration data and are excluded from Git. Delete the state file only when you intentionally want to reprocess the migration.

## Security

- Never commit `.env`, client secrets, access tokens, or migration state files.
- `.env` is ignored by `.gitignore`; `.env.example` contains placeholders only.
- Store client secrets in a local secret manager or local environment where possible.
- Rotate any secret that has been exposed, including in terminal output, chat, screenshots, or source control.
- Review the target folder and permissions before removing or changing source data. This tool copies data; it does not delete source files.

## Troubleshooting

### `Missing required environment variable`

Confirm that `.env` exists in the directory where the command is run and that all six tenant/app credential variables are populated.

### `itemNotFound` for a folder path

Check spelling and capitalization. Use paths relative to the OneDrive root and omit `My Files`.

### `accessDenied` or authentication errors

Confirm the correct tenant ID, client ID, and client secret are used for each app. Verify that the required application permission has admin consent in the matching tenant.

### `NotOpenSSLWarning` on macOS

This warning comes from an older Python build linked against LibreSSL. It is not the cause of a Graph API failure when the script continues. Use a current Python installation linked against OpenSSL if the warning needs to be removed.

### Throttling or temporary Graph failures

The script retries HTTP `429`, `500`, `502`, `503`, and `504` responses. For persistent failures, rerun the command; completed items remain recorded in the state file.

## Files

| File | Purpose | Check in? |
| --- | --- | --- |
| `migrate_onedrive.py` | Migration implementation | Yes |
| `requirements.txt` | Python dependencies | Yes |
| `.env.example` | Safe configuration template | Yes |
| `.env` | Local credentials and paths | No |
| `migration-state.json` | Local resume checkpoint | No |
| `.venv/` | Local Python environment | No |
