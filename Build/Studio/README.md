# Studio distribution

PowerForge owns Studio publishing, signing, checksums, portable archives, and the generated Windows MSI. Repository-local files contain only OfficeIMO product identity and target choices.

```powershell
./Build/Studio/Build-Studio.ps1 -Validate
./Build/Studio/Build-Studio.ps1 -Plan
./Build/Studio/Build-Studio.ps1 -Runtime win-x64
```

The release matrix contains self-contained `win-x64`, `win-arm64`, `osx-x64`, `osx-arm64`, `linux-x64`, and `linux-arm64` archives. The generated MSI uses a stable upgrade code and installs the `win-x64` build with a Start menu shortcut. User preferences and privacy-safe diagnostics remain under the user profile and are intentionally retained during ordinary uninstall.

Windows binaries and the MSI use the existing OfficeIMO Authenticode certificate profile and a trusted timestamp. A missing signing tool, certificate, or timestamp is a release failure. Do not disable signing for a public artifact.

Updates are manual for the initial product channel: install a newer signed artifact over the existing identity. Automatic update checks and Microsoft Store/App Installer publication remain disabled until a stable release feed and rollback policy exist. Building artifacts does not publish them.

The macOS target currently proves the self-contained runtime payload, not a notarized `.app`/`.dmg`. Native bundle construction, hardened-runtime signing, notarization, and stapling belong in reusable PowerForge desktop packaging before Studio consumes them. Linux initially ships as a checksum-verified portable archive; AppImage, Flatpak, and Debian/RPM packaging likewise belong in PowerForge, never in Studio-local scripts.
