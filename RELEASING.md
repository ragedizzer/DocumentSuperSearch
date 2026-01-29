# Release Checklist

## Before you release
- Confirm CI is green on main.
- Update the version header comments in:
  - Search.ps1
  - Search-Gui.ps1
  - Create-SearchGui-Shortcut.ps1
  - Test-ComObjects.ps1
- Run Test-ComObjects.ps1 on a target Windows machine.
- Run a small GUI search and a CLI search to verify output files and email.
- Update README.md if options or behavior changed.

## Release steps
- Merge via pull request into main (no direct pushes).
- Tag the release (for example: v1.2026).
- Draft release notes that summarize changes and known limitations.

## Branch protection expectations
- In the repository settings, protect the main branch with:
  - Require a pull request before merging.
  - Require at least one approving review.
  - Require status checks to pass: "CI / pwsh-parse".
  - Require conversation resolution and a linear history.
  - Disable force pushes and branch deletions.
