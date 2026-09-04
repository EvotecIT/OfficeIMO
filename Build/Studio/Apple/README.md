# Apple distribution channels

OfficeIMO Studio keeps direct macOS distribution and the Mac App Store as separate product channels. Users can choose either channel; Store preparation must not remove the signed and notarized direct download.

## Direct download

The active `powerforge.dotnetpublish.json` lane produces architecture-specific, multi-file self-contained `.app` bundles and `ditto` ZIP archives. PowerForge signs the native libraries in place instead of relying on single-file extraction. `Direct.entitlements` grants only the JIT permission required by the current non-NativeAOT .NET runtime.

Local proof uses explicit ad-hoc signing. A public artifact requires all of the following on a trusted macOS builder:

1. A `Developer ID Application` identity replaces the ad-hoc identity.
2. Secure timestamps remain enabled.
3. PowerForge submits the exact signed artifact for notarization, staples the ticket, and verifies it with `codesign` and Gatekeeper.
4. The release record binds the source commit, artifact SHA-256, signing identity, and notarization result.

## Mac App Store

`AppStore.entitlements` is a prepared sandbox profile, not an active release lane. It permits user-selected document read/write, outbound connections for approved online operations, and the JIT permission required by the current runtime.

The Store lane remains blocked until the shared PowerForge owner can package externally built macOS apps without an application-local script. That owner must embed a provisioning profile when the selected capabilities require one, sign with an Apple distribution identity, create and validate the installer package with a Mac installer distribution identity, and upload the exact package to App Store Connect. The builder also needs an App Store Connect app record for `com.evotec.officeimo.studio` and the corresponding credentials. These identities, profiles, and credentials remain outside the repository.

The App Sandbox changes product capabilities. Studio currently discovers and starts external tools such as Tesseract, LibreOffice, and Pandoc. A Store build must not assume it can execute arbitrary user-installed binaries. Each feature must instead use a permitted bundled and signed helper or be disabled with a contextual explanation and a direct-download alternative. File access must flow through user-selected URLs and retained security-scoped access where a later session needs the same document. Store builds use App Store updates; they do not run a parallel self-updater.

Before submission, validate receipt handling, container paths, privacy disclosures, accessibility, localization screenshots, clean install/update/uninstall behavior, and the complete rendered state matrix on a Store-signed build.

## Future iPhone and iPad apps

An iOS or iPadOS product is not another runtime identifier for the desktop executable. It should reuse the OfficeIMO document engines, workflow contracts, preferences/localization abstractions, and portable view models where appropriate, while owning a mobile interaction shell, document-picker and security-scoped storage adapters, lifecycle behavior, and platform-specific packaging in an Apple target.

App Store Connect can add future platforms to one app record when a shared product identity and universal purchase are intentional. That is a product decision, not a packaging default; a separately positioned mobile product can use its own record and bundle identifier.
