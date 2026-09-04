# Studio distribution

PowerForge owns Studio publishing, signing, checksums, portable archives, the generated Windows MSI, Debian package, and macOS application bundle. Repository-local files contain only OfficeIMO product identity and target choices.

The checked-in `packages.lock.json` files cover Studio's complete project-reference graph and all supported runtime identifiers. Release restores run in locked mode so package or runtime-asset changes must be reviewed and committed before signing.

Avalonia and CommunityToolkit.Mvvm are the only explicitly trusted build-code providers. PowerForge still verifies their exact archives through the committed lock before allowing the XAML compiler and MVVM source generator to execute; Avalonia's separate telemetry build package is excluded from Studio.

```powershell
./Build/Studio/Build-Studio.ps1 -Validate
./Build/Studio/Build-Studio.ps1 -Plan
./Build/Studio/Build-Studio.ps1 -Target Studio.Windows -Runtime win-x64
./Build/Studio/Build-Studio.ps1 -Target Studio.macOS -Runtime osx-arm64
./Build/Studio/Build-Studio.ps1 -Target Studio.Linux -Runtime linux-x64
```

Always pair a target with its compatible runtime when narrowing the release matrix. Do not use `-SkipBuild`: each runtime needs its own project-reference outputs before PowerForge performs the no-build publish and packages the result.

The release matrix contains self-contained `win-x64`, `win-arm64`, `osx-x64`, `osx-arm64`, `linux-x64`, and `linux-arm64` archives. The generated MSI uses a stable upgrade code and installs the `win-x64` build with a Start menu shortcut. The Debian package owns `/opt`, `/usr/bin`, freedesktop desktop metadata, MIME associations, and the application icon. The macOS package owns the `.app` layout, stable `com.evotec.officeimo.studio` bundle identifier, generated ICNS icon, document associations, code-signing verification, and a `ditto` ZIP. User preferences and privacy-safe diagnostics remain under the user profile and are intentionally retained during ordinary uninstall.

Windows binaries and the MSI use the existing OfficeIMO Authenticode certificate profile and a trusted timestamp. A missing signing tool, certificate, or timestamp is a release failure. Do not disable signing for a public artifact.

Updates are manual for the initial product channel: install a newer signed artifact over the existing identity. Automatic update checks and Microsoft Store/App Installer publication remain disabled until a stable release feed and rollback policy exist. Building artifacts does not publish them.

The checked-in macOS package uses explicit ad-hoc signing (`CodesignIdentity = "-"`) for local bundle and launch proof. That output is not a public direct-distribution artifact. Before public macOS distribution, set a `Developer ID Application` identity, enable trusted timestamps, and run the resulting exact `.app` through PowerForge's notarization, stapling, and Gatekeeper assessment flow. The trusted macOS builder must hold the signing identity and notary credentials; never put credentials in this configuration.

DMG, Linux AppImage/Flatpak/RPM, and any future native format belong in reusable PowerForge packaging rather than Studio-local scripts. The portable archives, MSI, Debian package, and macOS app ZIP are build outputs only; the build does not publish them.
