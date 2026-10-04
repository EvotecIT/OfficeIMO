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

The downloads page links signed Windows Studio MSIs to `downloads.officeimo.com`.
Evotec Control manages the release lifecycle and download activity through the
shared Licensing distribution service. The branded hostname serves the immutable
release objects directly over HTTPS.

[`Website/data/studio_installer_downloads.json`](../Website/data/studio_installer_downloads.json)
maps verified GitHub release asset URLs to their branded download URLs and records
the exact sizes and SHA-256 values. The downloads page uses the branded link only
when the release asset matches an entry. Portable archives, Linux packages and
other assets keep their GitHub release links. Installer bytes are held by the
distribution service and are not included in the website build.

When adding a release, publish and verify its signed artifacts through Control,
promote the exact release to public stable, and add its verified URL mapping.
Keep objects referenced by Microsoft Store or other catalogs available at their
original versioned URLs. Never replace bytes at a published release path.

Before using a URL in a Store submission, retrieve it without following redirects
and confirm HTTP 200, exact length, SHA-256 and Authenticode signature against the
release. GitHub release URLs remain suitable for WinGet manifests.

## Editing guidance

- Edit authored content in `Website/content/`, `Website/data/`, `Website/site.json`, `Website/pipeline.json`, and theme files under `Website/themes/officeimo/`.
- Do not hand-edit `Website/_site/`, `Website/_temp/`, or copied PSWriteOffice example files.
- `Website/data/release-hub.json` and `Website/data/xrefmap.json` are generated artifacts. Keep intentional structural updates, but avoid committing timestamp-only churn.
