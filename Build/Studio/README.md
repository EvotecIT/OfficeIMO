# Studio distribution

PowerForge owns Studio publishing, signing, checksums, portable archives, the generated Windows MSI, Debian package, and macOS application bundle. Repository-local files contain only OfficeIMO product identity and target choices.

The checked-in `packages.lock.json` files cover Studio's complete project-reference graph and all supported runtime identifiers. Release restores run in locked mode so package or runtime-asset changes must be reviewed and committed before signing.

Windows releases use `packages.lock.json`; Unix releases use `packages.nonwindows.lock.json`.
Ordinary development restores use temporary locks under `obj` on both platforms.
The SDK baseline in `global.json` and CI is part of the release restore contract:
single-file publishing adds SDK-provided analyzer packages. Refresh both platform
lock sets when that baseline or the publish dependency graph changes. CI checks
the six-runtime, single-file restore graph on Windows and Ubuntu.

Release restores disable SDK offline library packs so platform locks refer to the
configured package feed. Distribution SDKs can ship rebuilt analyzer packages
under the same version with different hashes. Use a fresh package directory when
refreshing release locks if the local cache contains those rebuilt packages.

Avalonia and CommunityToolkit.Mvvm are the only explicitly trusted build-code providers. PowerForge still verifies their exact archives through the committed lock before allowing the XAML compiler and MVVM source generator to execute; Avalonia's separate telemetry build package is excluded from Studio.

```powershell
./Build/Studio/Build-Studio.ps1 -Validate
./Build/Studio/Build-Studio.ps1 -Plan
./Build/Studio/Build-Studio.ps1 -Target Studio.Windows -Runtime win-x64
./Build/Studio/Build-Studio.ps1 -Target Studio.macOS -Runtime osx-arm64
./Build/Studio/Build-Studio.ps1 -Target Studio.Linux -Runtime linux-x64
./Build/Studio/Build-StudioWindowsRelease.ps1 -Validate
./Build/Studio/Build-StudioWindowsRelease.ps1 -Plan
./Build/Studio/Build-StudioLinuxRelease.ps1 -Validate
./Build/Studio/Build-StudioLinuxRelease.ps1 -Plan
```

Always pair a target with its compatible runtime when narrowing the release matrix. Do not use `-SkipBuild`: each runtime needs its own project-reference outputs before PowerForge performs the no-build publish and packages the result.

The release matrix contains self-contained `win-x64`, `win-arm64`, `osx-x64`, `osx-arm64`, `linux-x64`, and `linux-arm64` archives. Windows creates MSI installers for x64 and Arm64 with the same stable upgrade code and a Start menu shortcut. The Debian package owns `/opt`, `/usr/bin`, freedesktop desktop metadata, MIME associations, and the application icon. The macOS package owns the `.app` layout, stable `com.evotec.officeimo.studio` bundle identifier, generated ICNS icon, document associations, code-signing verification, and a `ditto` ZIP. User preferences and privacy-safe diagnostics remain under the user profile and are intentionally retained during ordinary uninstall.

Studio has its own `0.1.x` release line. The Studio project version, Debian package version, macOS bundle version, and Windows MSI version must agree for a direct-download release. Both release wrappers pass the committed project version to PowerForge, which rejects a mismatched native installer version. PowerForge resolves one monotonic version for both Windows MSIs and portable ZIPs, applies it to the application binaries, and reserves it through `studio-msi/officeimo-studio` Git tags when building a release. If a previous reservation has advanced the MSI version, update and commit the Studio and native package versions before another release attempt; the wrapper fails before building mismatched assets. NuGet library versions remain independent. Planning is read only; an actual Windows build with this release config reserves the version remotely, even when signing is disabled. For unsigned local package tests, use an isolated temporary copy of the config with signing disabled and the version authority changed to `LocalFile` with a task-owned state path. Never publish those test artifacts.

`Build-StudioWindowsRelease.ps1` stages the two signed MSI files, two portable ZIPs, Windows checksums and release manifest, and the three-file WinGet manifest set. For a release candidate, record the checkout's exact commit, then run the wrapper once with `-Publish` from a clean checkout: PowerForge binds the binaries and draft `Studio-v<version>` GitHub tag to that commit and uploads those exact bytes in the same build. Running the wrapper a second time builds new bytes and reserves a new version; it is not a way to publish the qualified candidate.

`Build-StudioLinuxRelease.ps1` stages the x64 and Arm64 Debian packages and portable ZIPs without publishing them. Run it from a clean Linux checkout at the **same exact commit** as the Windows draft. Compare its staged hashes and package metadata to `linux-release-manifest.json` and `linux-SHA256SUMS.txt`. Once both architectures pass installation and launch checks, verify that the GitHub draft tag points to that commit, then upload the four staged Linux packages plus those two Linux metadata files to that draft. The Windows and Linux manifests each cover their own platform; both must be present and verified. Download all eight draft packages and compare their hashes to the staged files before promoting the release. The website shows each platform only when both installer and portable assets for both architectures exist in the published release.

```powershell
$version = '<qualified Studio version>'
$sourceCommit = (git rev-parse HEAD).Trim()
$draft = gh release view "Studio-v$version" -R EvotecIT/OfficeIMO --json isDraft,targetCommitish | ConvertFrom-Json
if (-not $draft.isDraft -or $draft.targetCommitish -ne $sourceCommit) { throw 'The Windows draft does not match this Linux source commit.' }
$linuxRoot = 'Artifacts/Studio/LinuxRelease'
$linuxAssets = @(
    "$linuxRoot/GitHub/officeimo-studio_${version}_amd64.deb"
    "$linuxRoot/GitHub/officeimo-studio_${version}_arm64.deb"
    "$linuxRoot/GitHub/OfficeIMO-Studio-linux-x64-PortableCompat.zip"
    "$linuxRoot/GitHub/OfficeIMO-Studio-linux-arm64-PortableCompat.zip"
    "$linuxRoot/linux-release-manifest.json"
    "$linuxRoot/linux-SHA256SUMS.txt"
)
foreach ($asset in $linuxAssets) { if (-not (Test-Path -LiteralPath $asset)) { throw "Missing Linux release asset: $asset" } }
# Verify package metadata, architecture, executable permissions, hashes, and launch before uploading.
gh release upload "Studio-v$version" @linuxAssets -R EvotecIT/OfficeIMO
```

The Windows release wrapper requires public PSPublishModule 3.0.146 or later. The Linux release wrapper requires PSPublishModule 3.0.152 or later for native installer release-asset staging. An older build may produce a Debian package without including it in the staged release assets.

Before promoting, compare downloaded asset hashes to the staged checksums, verify Authenticode and timestamp evidence, and exercise Windows install, upgrade, launch, and uninstall on x64 and Arm64. Exercise Linux install, launch, and uninstall on x64 and Arm64. Confirm the draft contains all eight matching Windows and Linux packages, both platform manifests and checksum files, and that its `targetCommitish` equals the source commit recorded at build time. Only after that acceptance, publish the same release with `gh release edit "Studio-v$version" -R EvotecIT/OfficeIMO --target $sourceCommit --draft=false`. This step changes the release visibility without rebuilding artifacts. Verify the resulting tag ref. WinGet submission and Store distribution are separate channels.

```powershell
$version = '<qualified Studio version>'
$sourceCommit = '<exact commit recorded before the release build>'
$draft = gh release view "Studio-v$version" -R EvotecIT/OfficeIMO --json isDraft,targetCommitish,assets | ConvertFrom-Json
if (-not $draft.isDraft -or $draft.targetCommitish -ne $sourceCommit) { throw 'Studio draft does not match the qualified source commit.' }
# Complete the downloaded-asset qualification described above before the next command.
gh release edit "Studio-v$version" -R EvotecIT/OfficeIMO --target $sourceCommit --draft=false
$tagCommit = gh api "repos/EvotecIT/OfficeIMO/git/ref/tags/Studio-v$version" --jq '.object.sha'
if ($tagCommit -ne $sourceCommit) { throw 'Published Studio tag does not match the qualified source commit.' }
gh release view "Studio-v$version" -R EvotecIT/OfficeIMO --json isDraft,assets
```

The website reads published Studio assets from the release hub. It shows Windows MSI and portable links, Linux Debian and portable links, or macOS application ZIP links only when the complete asset set for that platform and both architectures is present in a published release. Otherwise it points to the source build guide. Before a public release, verify exact asset hashes and applicable signatures, then test clean installation, upgrade, uninstall, and launch on supported systems. WinGet catalog acceptance is a separate public-state check.

Windows binaries and the MSI use the existing OfficeIMO Authenticode certificate profile and a trusted timestamp. A missing signing tool, certificate, or timestamp is a release failure. Do not disable signing for a public artifact.

Updates are manual for the initial product channel: install a newer signed artifact over the existing identity. Automatic update checks and Microsoft Store/App Installer publication remain disabled until a stable release feed and rollback policy exist. Building artifacts does not publish them.

The checked-in macOS package uses explicit ad-hoc signing (`CodesignIdentity = "-"`) and the minimum JIT entitlement required by a non-NativeAOT .NET app for local bundle and launch proof. That output is not a public direct-distribution artifact. Before public macOS distribution, set a `Developer ID Application` identity, enable trusted timestamps, and run the resulting exact `.app` through PowerForge's notarization, stapling, and Gatekeeper assessment flow. The trusted macOS builder must hold the signing identity and notary credentials; never put credentials in this configuration.

Direct download and Mac App Store delivery are separate channels. The App Store channel uses its own sandbox profile, Apple distribution identities, store-owned updates, and a capability review for external helper processes. See [Apple/README.md](Apple/README.md) for the channel contract and future iOS boundary.

DMG, Linux AppImage/Flatpak/RPM, and any future native format belong in reusable PowerForge packaging rather than Studio-local scripts. The portable archives, MSI, Debian package, and macOS app ZIP are build outputs only; the build does not publish them.
