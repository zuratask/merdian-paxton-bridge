# NET2 Sync

PowerShell WinForms UI prototype for CSV drops, NET2 SQL checks, and renewals (mock data only; no real database yet).

## Run

```powershell
powershell -ExecutionPolicy Bypass -File ".\N2SYNC-UI.ps1"
```

Requires Windows with PowerShell 5.1+ and .NET Framework (WinForms).

## Publish to GitHub (on a machine with Git)

1. Install [Git for Windows](https://git-scm.com/download/win) if needed.
2. On GitHub: **New repository** (e.g. `net2-sync`), leave “Initialize with README” unchecked if you already have this folder.
3. In this directory:

```powershell
cd "c:\Users\Administrator\Downloads\N2SYNC"
git init
git add .
git commit -m "Initial commit: NET2 Sync UI prototype"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

Replace `YOUR_USERNAME` / `YOUR_REPO` with your GitHub user and repository name. Use a [personal access token](https://github.com/settings/tokens) as the password when Git prompts for credentials, or sign in with GitHub CLI (`gh auth login`).
