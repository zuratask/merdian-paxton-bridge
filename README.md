# NET2 Sync

PowerShell WinForms UI prototype for CSV drops, NET2 SQL checks, and renewals (mock data only; no real database yet).

## Run

```powershell
powershell -ExecutionPolicy Bypass -File ".\N2SYNC-UI.ps1"
```

Requires Windows with PowerShell 5.1+ and .NET Framework (WinForms).

On older Windows, if downloads fail with SSL errors, run once in PowerShell before `Invoke-WebRequest`:

```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
```

## Git and GitHub CLI (this machine)

Git for Windows and GitHub CLI (`gh`) can be installed under:

- `C:\Program Files\Git\cmd\git.exe`
- `C:\Program Files\GitHub CLI\gh.exe`

This repo is already initialized with `main` and an initial commit. To **create the repository on GitHub and push** (one-time login):

```powershell
cd "c:\Users\Administrator\Downloads\N2SYNC"
gh auth login
gh repo create net2-sync --public --source=. --remote=origin --push
```

Change `net2-sync` if the name is taken. If the repo already exists on GitHub instead:

```powershell
gh auth login
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

Use HTTPS and a [personal access token](https://github.com/settings/tokens) when Git asks for a password, or complete `gh auth login` with the browser/device flow.
