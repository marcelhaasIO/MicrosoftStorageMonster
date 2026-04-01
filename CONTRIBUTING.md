# Contributing

Thanks for your interest in contributing to Microsoft Storage Monster!

## Ways to Contribute

- **Bug reports** — open an issue describing the error, your PowerShell version, and steps to reproduce
- **Feature requests** — open an issue with the use case and expected behavior
- **Pull requests** — fork the repo, make your changes, and open a PR against `main`

## Guidelines

- Keep the script dependency-free except for PnP.PowerShell
- Maintain compatibility with PowerShell 5.1 and PowerShell 7+
- New finding types should follow the existing `FindingType` pattern and be documented in the README
- Test against at least one real SharePoint Online tenant before submitting

## Development

No build step required — the script is a single `.ps1` file. Edit and run directly.

```powershell
# Test against a single site
.\StorageScan-Local.ps1 -SiteUrl "https://yourtenant.sharepoint.com/sites/TestSite" -OutputPath ".\test-output"
```
