# Merdian Paxton Bridge

CSV drops, Paxton NET2 checks, and renewals — PowerShell WinForms UI prototype (mock data only; no real database yet).

**Repository:** https://github.com/zuratask/merdian-paxton-bridge

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

Git for Windows and GitHub CLI (`gh`) are typically installed at:

- `C:\Program Files\Git\cmd\git.exe`
- `C:\Program Files\GitHub CLI\gh.exe`

To clone:

```powershell
git clone https://github.com/zuratask/merdian-paxton-bridge.git
```

To push after changes:

```powershell
cd merdian-paxton-bridge
git add -A
git commit -m "Your message"
git push
```

If you use a new machine, sign in once with `gh auth login` or configure HTTPS credentials with a [personal access token](https://github.com/settings/tokens).
