# OfficeIMO Website

The OfficeIMO public website lives under [`Website/`](../Website/) and is built with PowerForge.Web from the sibling [`PSPublishModule`](https://github.com/EvotecIT/PSPublishModule) repository.

## Local build

From the website folder:

```powershell
.\build.ps1
```

Recommended local setup when working on the full site pipeline:

```powershell
.\build.ps1 -CI -PowerForgeRoot C:\Support\GitHub\PSPublishModule -PSWriteOfficeRoot C:\Support\GitHub\PSWriteOffice
```

- `-PowerForgeRoot` points to the local PowerForge.Web engine checkout.
- `-PSWriteOfficeRoot` refreshes the PowerShell API snapshot from the sibling `PSWriteOffice` repo before building.

## API inputs

The website publishes two API surfaces.

### .NET package API

Generated from compiled XML docs and assemblies during the website pipeline.

The generated API surface includes the main document packages and their optional integrations. OCR is published as five separate API routes so consumers can see the neutral contract, providers, and format adapters independently:

- `OfficeIMO.Word`
- `OfficeIMO.Excel`
- `OfficeIMO.Markdown`
- `OfficeIMO.PowerPoint`
- `OfficeIMO.CSV`
- `OfficeIMO.Visio`
- `OfficeIMO.Reader`
- `OfficeIMO.Ocr`
- `OfficeIMO.Ocr.Process`
- `OfficeIMO.Ocr.Tesseract`
- `OfficeIMO.Reader.Ocr`
- `OfficeIMO.Pdf.Ocr`

[`Website/pipeline.json`](../Website/pipeline.json) is the complete source of truth for every generated API route.

The merged cross-reference map is committed at [`Website/data/xrefmap.json`](../Website/data/xrefmap.json).

### PSWriteOffice PowerShell API

Generated from a synced `PSWriteOffice` repo snapshot when available, with checked-in website fallback inputs under:

- [`Website/data/apidocs/powershell/PSWriteOffice-Help.xml`](../Website/data/apidocs/powershell/PSWriteOffice-Help.xml)
- [`Website/data/apidocs/powershell/examples/`](../Website/data/apidocs/powershell/examples/)

Those inputs are refreshed by [`Website/scripts/Sync-PSWriteOfficeApiDocs.ps1`](../Website/scripts/Sync-PSWriteOfficeApiDocs.ps1).

The sync script:

- looks for a synced or local `PSWriteOffice` repo
- prefers `Docs/Generated/PSWriteOffice-help.xml` from that repo
- falls back to `Artefacts/Unpacked/Modules/PSWriteOffice/en-US/PSWriteOffice-help.xml`
- mirrors the repo `Examples/` folder into the website fallback folder
- preserves clean-checkout behavior when the source repo is unavailable

## CI / GitHub Pages

Website automation lives in:

- [`.github/workflows/website-ci.yml`](../.github/workflows/website-ci.yml)
- [`.github/workflows/deploy-website.yml`](../.github/workflows/deploy-website.yml)

Both workflows:

- check out `PSPublishModule`
- run `sources-sync`, which pulls `EvotecIT/PSWriteOffice` into `Website/projects/pswriteoffice`
- run the PowerShell API sync script
- build the website in CI mode
- verify required output routes before upload/deploy

If the synced repo does not contain the generated help snapshot, the build falls back to the checked-in PowerShell API inputs instead of failing.

## Studio installer downloads

The website publishes signed Studio MSIs at immutable versioned paths such as
`https://officeimo.com/downloads/studio/0.1.9767/OfficeIMO-Studio-0.1.9767-win-x64.msi`.
These files are included in the GitHub Pages deployment so the URL serves the
installer directly, without a redirect to GitHub's release storage.

[`Website/data/studio_installer_downloads.json`](../Website/data/studio_installer_downloads.json)
owns the source URLs, relative destination paths, exact sizes and SHA-256 pins.
PowerForge's `download-artifacts` pipeline step checks every pin after the final
site build. The downloads page uses local links only for matching pinned release
assets; portable archives and other release assets keep their release links.

When adding a release, verify its signed artifacts and append new versioned
entries. Keep entries used by Microsoft Store or other catalogs so later website
deployments continue to publish those URLs. Never replace the bytes at a published
versioned path. Check the complete deployment against GitHub Pages' one-gigabyte
site limit; the manifest's byte limit covers installer downloads only.

The site audit exempts the two pinned MSI paths from its general 25 MiB file
budget. Their exact sizes and hashes are enforced by `download-artifacts`.
Add the corresponding exact audit exception when pinning a new installer; keep
the general budget for other site assets.

Before using a URL in a Store submission, retrieve it without following redirects
and confirm HTTP 200, exact length, SHA-256 and Authenticode signature against the
release. GitHub release URLs remain suitable for WinGet manifests.

## Editing guidance

- Edit authored content in `Website/content/`, `Website/data/`, `Website/site.json`, `Website/pipeline.json`, and theme files under `Website/themes/officeimo/`.
- Do not hand-edit `Website/_site/`, `Website/_temp/`, or copied PSWriteOffice example files.
- `Website/data/release-hub.json` and `Website/data/xrefmap.json` are generated artifacts. Keep intentional structural updates, but avoid committing timestamp-only churn.
