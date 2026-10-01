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
| Country availability | All current App Store countries and regions, including new regions added by Apple |

iPadOS uses the iOS store entry. The Mac entry is a separate native application
binary under the same record. An iPhone app running on Apple silicon does not
qualify the native Mac product.

`powerforge.store-onboarding.json` owns remote setup without build targets.
`powerforge.release.json` selects the native .NET Mac App Store archive target;
its `DotNetPublishInstallerId` routes through PowerForge's shared MacApp packager.
The selected `powerforge.dotnetpublish.json` builds a self-contained Apple silicon
single-file host with native libraries kept outside the executable for signing.
No GUI adaptation or mobile placeholder host is part of this configuration.

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

`AppStore.entitlements` enables the App Sandbox, user-selected document read/write,
app-scoped bookmarks, outbound connections, and JIT support. Store builds use
`OfficeIMOStudioDistribution=MacAppStore`: app data stays in the Apple container,
external Tesseract discovery and printer command execution are blocked with a
contextual explanation. In-process document conversions and print-PDF preparation
remain available. Document conversions do not require external office
applications or converters.

The shared packager places the host and native code in `Contents/MacOS`, notices
and data in `Contents/Resources`, signs nested Mach-O code before the app, and
creates a signed `.pkg` with `productbuild`. It checks the app's distribution
team, installer signature, and the expanded installer payload. Runtime code must
use bundle resource paths for external content; Store data is not beside the host.

Use the PowerForge CLI built from the shared source containing this backend:

```text
powerforge apple-release Archive --config Build/Studio/Apple/powerforge.release.json --plan --summary --output json
powerforge apple-release Archive --config Build/Studio/Apple/powerforge.release.json --summary --output json
```

The second command creates a signed package and `.xcarchive` locally. It requires
an Apple distribution application identity and a Mac installer distribution
identity for team `8ZPGZ79T7J`. A development certificate can qualify sandbox
behavior but cannot produce a Store installer. Install credentials on the trusted
builder; never commit private keys, certificates, keychains, or provisioning
profiles. If capabilities require a profile, set `ProvisioningProfilePath` to a
builder-local path within the configured project root. The packager rejects
expired, wrong-team, wrong-bundle and development profiles.

Keep marketing version and build number identical in the release target and
MacApp configuration. This backend uses explicit versions; Xcode project
generation, automatic version mutation, and the Swift exact-package snapshot
mode do not apply. Archive creation is the supported local preparation boundary;
export/upload and Store ingestion need qualification with distribution identities.
Upload, metadata synchronization, review submission and release are disabled.

Before submission, qualify user-selected open/save, retained bookmarks after
relaunch, recovery and recent documents, network features, clean install/update,
and accessibility on the final Store-signed build. Capture screenshots from that
build. Store builds use Apple's update channel.

The privacy policy is maintained in `OfficeIMO.Studio/PRIVACY.md` and the public
website page `Website/content/pages/studio-privacy.md`. Publish the policy before
syncing its URL. Local document processing, container diagnostics and settings
must be distinguished from optional remote assistance: the selected provider
receives questions and document evidence, and its account and retention terms
apply. Review provider behavior and Apple's optional-collection exceptions before
answering the privacy questionnaire; do not infer a blanket "data not collected"
answer from local processing.

Apple's [required-reason API guidance](https://developer.apple.com/documentation/bundleresources/describing-use-of-required-reason-api)
covers iOS, iPadOS, tvOS, visionOS and watchOS. Assess the actual mobile host and
its SDKs before declaring reasons. Studio uses file timestamps for container
recovery/signatures and selected-document thumbnail invalidation. Avalonia's
native macOS storage provider creates and resolves security-scoped bookmarks.
These observations are inputs to qualification, not a completed mobile manifest.

Studio's protected-PDF features use application-level cryptography, including
BouncyCastle. Assess that exact graph for export compliance; do not set
`ITSAppUsesNonExemptEncryption=false` merely because network requests use HTTPS.
Derive required-reason privacy manifests from the final runtime and native
dependencies. An empty manifest or an unverified "data not collected" answer is
not a privacy audit, especially when remote document assistance is enabled.

## Open-source distribution

The app publishes the repository MIT license, WebM/libvpx license and patent grant,
sRGB profile terms, and
[the open-source explanation](../../../OfficeIMO.Studio/OPEN_SOURCE.md) under
`Licenses/`. Store packages retain these files in
`Contents/Resources/Licenses/`, alongside generated `THIRD_PARTY_NOTICES.txt` and
`runtime-package-inventory.json`. `third-party-notices.json` binds reviewed
license text hashes to exact published package versions. Dependency upgrades
require refreshed coverage; unknown versions fail packaging.
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
