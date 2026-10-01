# Apple distribution channels

OfficeIMO Studio keeps direct macOS distribution and the Mac App Store as separate product channels. Users can choose either channel; Store preparation must not remove the signed and notarized direct download.

## Store identity and onboarding

OfficeIMO Studio uses one App Store Connect record for the native macOS app and
the iPhone/iPad app:

| Setting | Value |
| --- | --- |
| Name | OfficeIMO Studio |
| App Store Connect app ID | `6818096229` |
| Bundle ID | `com.evotec.officeimo.studio` |
| Apple team | `8ZPGZ79T7J` |
| SKU | `officeimo-studio` |
| Primary language | English (U.S.) |
| Store price | Free |

iPadOS uses the iOS store entry. The Mac entry is a separate native application
binary under the same record. An iPhone app running on Apple silicon does not
qualify the native Mac product.

`powerforge.store-onboarding.json` wires the governance and metadata files for
onboarding. It has no archive targets: PowerForge's current Apple archive lane
requires Xcode projects, while Studio is built with .NET. `apple-release Doctor`,
`Advance`, and submission require an implemented archive backend and actual
targets; an empty `Apps` array is not a shipping release configuration.

Load the existing App Store Connect environment from the trusted builder, then
use the shared engine for the supported governance operations:

```text
powerforge apple-governance validate --config Build/Studio/Apple/appstoreconnect-governance.json
powerforge apple-governance plan --config Build/Studio/Apple/appstoreconnect-governance.json --release-config Build/Studio/Apple/powerforge.store-onboarding.json --receipt Artifacts/Studio/Apple/governance-plan.json --fail-on-drift --summary --output json
```

Apply a reviewed full plan with `apple-governance apply --reviewed-plan ...
--confirm`, then replan. The committed declaration owns only its populated
sections. Empty accessibility, encryption, and subscription arrays carry no
claims and do not erase remote resources. Availability is declared separately
from price; a free price does not make an unpublished app available.

The app-information and macOS metadata files are submission inputs, not proof
of a qualified Store build. Publish the linked privacy policy before syncing its
URL. Check every feature against the final sandbox build before syncing the
description. Mobile descriptions and screenshots must come from the mobile
product rather than desktop capabilities. Upload, review, and release actions
remain disabled in the onboarding configuration.

Open implementation and qualification work belongs in
[the Studio roadmap](../../../Docs/ROADMAP.md#desktop-studio).

## Direct download

The active `powerforge.dotnetpublish.json` lane produces architecture-specific, multi-file self-contained `.app` bundles and `ditto` ZIP archives. PowerForge signs the native libraries in place instead of relying on single-file extraction. `Direct.entitlements` grants only the JIT permission required by the current non-NativeAOT .NET runtime.

Local proof uses explicit ad-hoc signing. A public artifact requires all of the following on a trusted macOS builder:

1. A `Developer ID Application` identity replaces the ad-hoc identity.
2. Secure timestamps remain enabled.
3. PowerForge submits the exact signed artifact for notarization, staples the ticket, and verifies it with `codesign` and Gatekeeper.
4. The release record binds the source commit, artifact SHA-256, signing identity, and notarization result.

## Mac App Store

`AppStore.entitlements` is a prepared sandbox profile, not an active release lane. It permits user-selected document read/write, outbound connections for approved online operations, and the JIT permission required by the current runtime.

The Store lane remains blocked until the shared PowerForge owner can package externally built macOS apps without an application-local script. That owner must embed a provisioning profile when the selected capabilities require one, sign with an Apple distribution identity, create and validate the installer package with a Mac installer distribution identity, and upload the exact package to the record above. These identities, profiles, and credentials remain outside the repository. In particular, never commit App Store Connect API private keys (`AuthKey_*.p8`), signing certificates or private-key bundles (`*.p12` or `*.pfx`), provisioning profiles, keychains, passwords, or authentication exports. Store them only in the trusted builder's secret and keychain facilities.

The App Sandbox changes product capabilities. Studio currently discovers and starts external tools such as Tesseract, LibreOffice, and Pandoc. A Store build must not assume it can execute arbitrary user-installed binaries. Each feature must instead use a permitted bundled and signed helper or be disabled with a contextual explanation and a direct-download alternative. File access must flow through user-selected URLs and retained security-scoped access where a later session needs the same document. Store builds use App Store updates; they do not run a parallel self-updater.

Before submission, validate receipt handling, container paths, privacy disclosures, accessibility, localization screenshots, clean install/update/uninstall behavior, and the complete rendered state matrix on a Store-signed build.

Studio's protected-PDF features use application-level cryptography, including
BouncyCastle. Assess that exact graph for export compliance; do not set
`ITSAppUsesNonExemptEncryption=false` merely because network requests use HTTPS.
Derive required-reason privacy manifests from the final runtime and native
dependencies. An empty manifest or an unverified "data not collected" answer is
not a privacy audit, especially when remote document assistance is enabled.

## Open-source distribution

The app publishes the repository MIT license, the WebM/libvpx notice, and
[the open-source explanation](../../../OfficeIMO.Studio/OPEN_SOURCE.md) under
`Licenses/`. The macOS bundler retains published files in
`Contents/MacOS/Licenses/`. Preserve these files when producing store packages.
Review every managed and native dependency in the exact published artifact and
include its required notices before distribution. The website dependency
inventory identifies packages; it does not prove notice coverage for a binary.

The first-party software stays MIT licensed. The free store price is a product
choice, separate from the license. Do not add a store-only source license,
subscription tier, license server, or source-disclosure requirement. Apple’s
standard store terms do not replace notices for the included open-source code.

## Future iPhone and iPad apps

An iOS or iPadOS product is not another runtime identifier for the desktop executable. It should reuse the OfficeIMO document engines, workflow contracts, preferences/localization abstractions, and portable view models where appropriate, while owning a mobile interaction shell, document-picker and security-scoped storage adapters, lifecycle behavior, and platform-specific packaging in an Apple target.

The iPhone/iPad host uses `com.evotec.officeimo.studio` and the existing iOS
record. It needs a real .NET iOS application host and device-family declaration,
AOT-compatible engine dependencies, document-picker and permission adapters,
and signed physical-device proof. Avoid a blank placeholder app or a Swift
shell that duplicates the C# document engines. GUI adaptation follows the mobile
feasibility criteria in the roadmap.
