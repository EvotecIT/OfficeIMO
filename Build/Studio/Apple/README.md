# Apple distribution

OfficeIMO Studio uses the native Swift app for the Mac App Store and the iPhone/iPad App Store. Both targets call the shared OfficeIMO C# document engine through its NativeAOT C interface. The current supported features are listed on the [Studio page](https://officeimo.com/studio/#features); the Apple release does not require the Windows/Linux interface or feature set.

## Store identity

| Setting | Value |
| --- | --- |
| Name | OfficeIMO Studio |
| App Store Connect app ID | `6818096229` |
| Bundle ID | `com.evotec.officeimo.studio` |
| Apple team | `8ZPGZ79T7J` |
| SKU | `officeimo-studio` |
| Primary language | English (U.S.) |
| Store price | Free |

The iOS target includes iPhone and iPad. The native Mac target is a separate binary under the same App Store Connect record. `powerforge.release.json` selects both Xcode schemes and reads their shared identity from `Apple/OfficeIMOStudio/Version.xcconfig`. Action flags stay false in Git; named PowerForge actions select the transition to execute.

## Engine artifact and notices

The local [Swift package](../../../Apple/OfficeIMOStudioEngine/Package.swift) pins the versioned XCFramework by URL and SHA-256. The engine stays MIT licensed and uses the existing OfficeIMO libraries. Its notice resource target carries the first-party license files and exact dependency inventories for all four native runtime slices.

Build an engine with `OfficeIMO.Studio.Native/Build-AppleEngine.ps1`, then package it with `Pack-StudioAppleEngine.ps1` as shown in the [native app guide](../../../Apple/OfficeIMOStudio/README.md#build). The packer uses PowerForge for dependency notice validation, archive creation, provenance metadata and optional GitHub release publication. `third-party-notices.json` binds reviewed texts to exact package versions; missing coverage fails packaging. Keep the artifact URL, checksum and generated package licenses together when updating the engine.

## Local native development

Use PowerForge's native deployment entrypoint with the release configuration:

```text
powerforge apple-deploy --config Build/Studio/Apple/powerforge.release.json --platform macOS --target "OfficeIMO Studio macOS"
powerforge apple-deploy --config Build/Studio/Apple/powerforge.release.json --platform iOS --target "OfficeIMO Studio iOS" --device-id <device-id>
```

The native Mac app uses the App Sandbox, user-selected document access and the system print panel. iPhone and iPad use the system document browser, photo picker and document camera. Document processing and OCR run on the device. These targets do not ship the Avalonia assistant or its remote-provider account flow.

## TestFlight and App Review

Load App Store Connect credentials from the trusted builder's private profile. Keep keys, provisioning profiles, worker policy and durable receipts outside Git. Use the same PowerForge release owner for both platforms:

```text
powerforge apple-release Doctor --config Build/Studio/Apple/powerforge.release.json --summary --output json
powerforge apple-release Version --config Build/Studio/Apple/powerforge.release.json --plan --summary --output json
powerforge apple-release Version --config Build/Studio/Apple/powerforge.release.json --confirm-apple-action --summary --output json
powerforge apple-release Advance --config Build/Studio/Apple/powerforge.release.json --plan --summary --output json
powerforge apple-release Advance --config Build/Studio/Apple/powerforge.release.json --confirm-apple-action --summary --output json
powerforge apple-release SubmitAppReview --config Build/Studio/Apple/powerforge.release.json --plan --summary --output json
powerforge apple-release SubmitAppReview --config Build/Studio/Apple/powerforge.release.json --confirm-apple-action --summary --output json
```

Commit the shared version change before archiving. Review each plan and use the exact-source release mode for signed candidates. `Advance` prepares builds, store metadata and the configured internal `Studio Internal` TestFlight group. Submission to public App Review is a separate action. The same Store-eligible binary can serve internal testing and App Review; do not export it as TestFlight Internal Only.

The two metadata files describe each native platform's supported scope. The two screenshot mappings select native Mac, iPhone and iPad PNGs under ignored `Artifacts/Studio/Apple/screenshots`. Keep those captures and their source-bound `approval-macOS.json` / `approval-iOS.json` manifests private until review. Screenshots must show the corresponding native product and use reviewed byte manifests. Qualify open/save, markup, page organization, form filling, scan/conversion, recovery and accessibility on the actual candidates. Simulator proof does not establish physical-device acceptance. Open work belongs in [the Studio roadmap](../../../Docs/ROADMAP.md#desktop-studio).

## Store declarations

`powerforge.store-onboarding.json` owns remote setup without archive targets. `appstoreconnect-governance.json` declares free pricing, worldwide availability and the reviewed encryption facts. Empty accessibility and subscription arrays make no claims and erase nothing. Plan and apply governance through PowerForge, retaining the exact reviewed plan.

The native privacy manifest records required-reason APIs used by document recovery, selected files, preferences and timing. The public policy is maintained in `OfficeIMO.Studio/PRIVACY.md` and [the website privacy page](https://officeimo.com/studio/privacy/). Review the native app's actual behavior when answering Apple's privacy questionnaire.

The bundled engine implements PDF encryption and certificate signatures. MIT/public source does not establish export exemption. Apple's declaration and any required documents must be accepted before the relevant review gate; do not infer readiness from a locally generated manifest.

## TestFlight automation

The **TestFlight** GitHub workflow provides **Build now**, **Automatic: Off**, **Automatic: Selected PRs**, and **Automatic: Every merged PR**. It records a request for the trusted PowerForge worker; GitHub does not receive Apple signing credentials. The worker qualifies the exact `master` commit with **Native Apple qualification** and deploys both native targets to the configured internal group. Automatic builds start disabled. Private worker enrollment and operator paths stay outside this public repository.

The **Studio Apple monitor** runs the shared read-only Doctor daily and retains its receipt. Its incident is updated only when the observed release state changes. The worker and monitor use the same pinned PowerForge source.

## Open-source distribution

The first-party software remains MIT licensed. Store pricing is a separate product choice. Apple distribution does not replace the copyright notices, permission grants or third-party license obligations. See [open-source notices](../../../OfficeIMO.Studio/OPEN_SOURCE.md).

## Avalonia Mac development

Use the certificate-signed development build for everyday file access:

```powershell
./Build/Studio/Build-StudioMacDevelopment.ps1 -Plan
./Build/Studio/Build-StudioMacDevelopment.ps1
```

The wrapper invokes PowerForge with `powerforge.development.json`. It builds the
Apple silicon app in Debug and signs it with the configured Apple Development
identity for team `8ZPGZ79T7J`. Managed assemblies stay inside the single-file
host; native libraries are signed separately, and notices live in
`Contents/Resources`. The bundle ID stays
`com.evotec.officeimo.studio`. PowerForge signs nested native libraries, verifies
the app and archive, and rejects a missing certificate or a certificate from the
wrong team or distribution channel. It never falls back to ad-hoc signing.

Use a PowerForge source build containing the `MacApp.DevelopmentOnly` option.
Set `POWERFORGE_SOURCE` to that checkout and build it with its pinned SDK:

```sh
(cd "$POWERFORGE_SOURCE" && dotnet build PowerForge.Cli/PowerForge.Cli.csproj -c Release -f net10.0)
export POWERFORGE_CLI="$POWERFORGE_SOURCE/PowerForge.Cli/bin/Release/net10.0/PowerForge.Cli.dll"
pwsh -File Build/Studio/Build-StudioMacDevelopment.ps1
```

The `.app` is at
`Artifacts/Studio/Development/packages/osx-arm64/OfficeIMO Studio.app`.
Quit Studio before replacing the installed copy in `/Applications`, and launch
that same copy for file-access checks. Keep the signing identity and bundle ID
consistent across rebuilds. Contributors use their own Apple Development identity
and team in their local configuration; no private keys belong in the repository.

macOS can require one fresh approval when moving from an ad-hoc build to this
identity. Later builds signed with the same identity retain a stable designated
requirement instead of identifying themselves by the executable hash. Running
`dotnet run` or an ad-hoc bundle does not exercise this signed app identity.
Development signatures are for local use; this build does not publish, notarize,
or create a Store submission. See Apple's
[code identity guidance](https://developer.apple.com/documentation/technotes/tn3127-inside-code-signing-requirements).

## Avalonia direct download

The [direct-download configuration](../powerforge.dotnetpublish.json) produces architecture-specific, multi-file self-contained `.app` bundles and `ditto` ZIP archives. PowerForge signs the native libraries in place instead of relying on single-file extraction. `Direct.entitlements` grants only the JIT permission required by the current non-NativeAOT .NET runtime.

The direct-download package configuration uses ad-hoc signing for packaging proof only. Use the development configuration above for routine local app use. PowerForge omits hardened runtime for ad-hoc bundles because their nested libraries do not share an Apple Team ID. The configured hardened-runtime option takes effect with a distribution signing identity; an ad-hoc launch does not qualify that contract.

A public artifact requires all of the following on a trusted macOS builder:

1. A `Developer ID Application` identity replaces the ad-hoc identity.
2. Secure timestamps remain enabled.
3. PowerForge submits the exact signed artifact for notarization, staples the ticket, and verifies it with `codesign` and Gatekeeper.
4. The release record binds the source commit, artifact SHA-256, signing identity, and notarization result.
