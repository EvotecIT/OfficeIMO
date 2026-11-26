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

## Engine and notices

The apps link the engine built from the same checkout, the way the rest of the app's code is built. `Build-StudioEngine.ps1` builds it with `OfficeIMO.Studio.Native/Build-AppleEngine.ps1` into the local [Swift package](../../../Apple/OfficeIMOStudioEngine/Package.swift) and skips the build while the engine sources are unchanged. An App Store build runs it before archiving, so a release always ships the engine from the commit it was archived from. Nothing is published to build or test the apps.

A four-slice NativeAOT build takes over an hour on GitHub's Mac runners, so pull requests don't build the engine. The Studio Apple engine workflow builds it from master each night when the engine sources changed. It keeps the build in the Actions cache and keeps one copy as the `studio-apple-engine` artifact, because the cache can be evicted. The Studio Apple job compiles the apps against the cached build, then the artifact, and only then the last released engine (`studio-engine-v1.0.0`). It warns when that engine lacks exports the native header declares. After a merge, a Studio Apple run on master that didn't get a current engine from the cache starts the engine workflow straight away. A pull request that changes the engine's C interface together with the Swift that calls it therefore needs a local build to check. CI checks it once the new engine is built from master.

The engine stays MIT licensed. The package's notice target carries the first-party license files and dependency notices. When the engine's dependencies change, `Pack-StudioAppleEngine.ps1` regenerates and validates the notices with PowerForge: `third-party-notices.json` binds reviewed texts to exact package versions, and missing coverage fails. Pass `-EngineDirectory Apple/OfficeIMOStudioEngine` after a full `Build-StudioEngine.ps1` run, which installs the dependency manifests for every runtime next to the framework.

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

Build the engine with `./Build/Studio/Apple/Build-StudioEngine.ps1` (all four slices) and commit the shared version change before archiving. Review each plan and use the exact-source release mode for signed candidates. `Advance` prepares builds, store metadata and the configured internal `Studio Internal` TestFlight group. Submission to public App Review is a separate action. The same Store-eligible binary can serve internal testing and App Review; do not export it as TestFlight Internal Only.

The metadata files describe each native platform's supported scope, in English (`appstoreconnect-metadata-*.json`, `appstoreconnect-app-info.json`) and in each other listing language under `appstore/localizations/<locale>/`. The screenshot mappings select the composed store art for each listing language under ignored `Artifacts/Studio/Apple/screenshots/store-art` (see [store screenshots in every language](#store-screenshots-in-every-language)). Keep those captures and their source-bound `approval-macOS.json` / `approval-iOS.json` manifests (`approval-<platform>.<locale>.json` for the other languages) private until review. Screenshots must show the corresponding native product and use reviewed byte manifests. Qualify open/save, markup, page organization, form filling, scan/conversion, recovery and accessibility on the actual candidates. Simulator proof does not establish physical-device acceptance. Open work belongs in [the Studio roadmap](../../../Docs/ROADMAP.md#desktop-studio).

The **Studio Apple screenshots** workflow captures a qualified source commit through PowerForge, pinned in the workflow and in `Invoke-PinnedPowerForge.ps1`. Select a verified trusted Apple silicon Mac runner and its available, Shutdown iPhone and iPad UUIDs from the default simulator device set. Their portrait display sizes must match the screenshot mappings above. The runner needs Xcode, PowerShell, the pinned .NET SDK and existing native UI access; automatic Mac window resizing also needs existing Accessibility access. The capture never grants permissions. Coordinate exclusive native UI use before dispatching it.

The capture-only XCTest targets build the Release app, open the repository's public samples through the normal document interface and retain native screenshots after recognizing each document's content and stable page pixels. The samples are the Northwind dashboard PDF, the equipment inspection form and the quarterly sales workbook from `Website/static/downloads/showcase`. The Mac host launches the exact built bundle with one public input at a time and the listing's language arguments; XCTest attaches to that process and checks its foreground identity. PowerForge owns simulator boot and verified shutdown, then retains the PNG inventory and source provenance. Review the actual retained images before creating approval manifests. Capturing screenshots does not approve or upload them to Apple.

The capture script installs the exact iOS app before copying the PDF. Its test manifest then uses `UseUITargetAppProvidedByTests` and excludes that app from Xcode's installation paths, preserving the document container. The pinned PowerForge process runner bounds Xcode execution and teardown so a failed test can return to app cleanup; native logs and validation receipts stay in the work folder. Mac launch arguments skip persistent document restoration; cleanup identifies the launched process by its PID, executable path and start time. The test accepts an operator-prepared 1280x800 window; automatic resizing requires existing Accessibility access for the test runner. Keep the document window unobstructed. The capture rejects other window edges crossing it and checks the actual document pixels. Full-screen transparent system surfaces appear in window metadata too, so geometry alone cannot prove that an image is unobstructed; inspect the exact PNGs before approval. The capture does not dismiss system notices or permission prompts.

Recognition and pixel stability use the native page's reported bounds clipped to the captured viewport. The heading, chart and lower decisions table must all be recognized in the same screenshot. Overlapping OCR crops preserve small table text. Mac queries stay inside the document window to avoid enumerating unrelated recent-document menu entries. The small Mac input uses native temporary storage and the app uses its own sandbox for scratch files; builds and diagnostics remain in the selected work folder.

### iPhone Duo

App Store Connect has two optional iPhone Duo screenshot sizes, required for iOS 27.1 SDK builds from April 2027: the outer display (1398×2034 pixels) and the inner display in portrait (2007×2853). Its API has no Duo display type yet, so the screenshot mappings leave them out and the sets are uploaded in App Store Connect by hand.

`simctl` can't fold the simulator, so Duo captures run locally, one pose per run, rather than in the workflow. Use an iPhone Duo simulator on iOS 27.1 that starts shut down, and a PowerForge build with `apple-simulator run --duo-pose` (the version CasaRay and Tactra pin, or later):

```sh
export POWERFORGE_DOTNET=<dotnet> POWERFORGE_SIMULATOR_CLI=<PowerForge.Cli.dll>
export POWERFORGE_CAPTURE_SIMULATORS='{"duo":"<iPhone Duo UUID>"}'
pwsh Build/Studio/Apple/Capture-StudioAppleScreenshots.ps1 -Role duo -DuoPose closed -WorkRoot <scratch>/studio-duo
pwsh Build/Studio/Apple/Capture-StudioAppleScreenshots.ps1 -Role duo -DuoPose open -WorkRoot <scratch>/studio-duo
```

Before the test starts, PowerForge checks that the pose's display is live at its exact size. A freshly booted Duo is closed. For the open pose, it asks on the terminal to open the device in Device Hub and turn it to portrait, or folds it with `POWERFORGE_DUO_FOLD_TOOL`, an opt-in community tool described in PowerForge's simulator guide. The Duo opens the PDF through the iPhone's document browser, whose Browse tab moves to a side rail on iOS 27, and captures the dashboard, markup and pages, in English, under `Duo/en-US/duo-outer-1398x2034` and `duo-inner-2007x2853`. The script checks each PNG's pixel size. Both runs can share one work folder, so the second reuses the first build.

### Store screenshots in every language

The store listing has seven languages, the same as the app's String Catalogs: en-US, de-DE, fr-FR, pl, es-ES, it and pt-BR. `store-art/locales.json` maps each listing locale to the app language and region a capture runs in.

The capture test drives the app through accessibility identifiers (`Shared/App/StudioAccessibilityID.swift`, compiled into the app and the capture bundles), never through translated labels, so one test covers every language. The document samples stay English; their text is what the readiness check recognizes. Each device role keeps these scenes:

| Role | Scenes |
| --- | --- |
| iPhone | dashboard, markup, form, Excel viewer, pages |
| iPad | dashboard, search, markup, form, Excel viewer, inspector |
| Mac | dashboard, search, markup, form, Excel viewer, inspector |

`Capture-StudioAppleScreenshots.ps1 -Role all` captures every locale by default; `-Locales de-DE,pl` selects some. For each locale it launches the Mac app with `-AppleLanguages` and `-AppleLocale`, and runs each simulator in its own PowerForge session after setting the shut-down device's system language and region, so the status bar matches too. The en-US session writes `build/powerforge/screenshots/<role>.json`, which the workflow verifies; other locales write `<role>-<locale>.json`. Afterwards the script returns both simulators to English. Raw PNGs land in `Artifacts/Studio/Apple/screenshots/<iPhone|iPad|macOS>/<locale>/`, each checked for its App Store pixel size and for visible content (`verify-screenshot-content.swift`).

The script then composes the store art with PowerForge's `apple-store-art compose` (theme `officeimo`, devices iPhone 6.9", iPad 13" and Mac at 2880×1800) from `store-art/store-art.json` and the captions in `store-art/captions/store-captions.<locale>.json`. Captions are written natively for each language, short and factual, with no pricing; have a native speaker review a changed caption before approval. Composed PNGs land in `screenshots/store-art/<locale>/<device>/`, and the workflow verifies them against the raw captures. A local composition needs a PSPublishModule checkout at the pinned commit, beside this repository or in `PSPUBLISHMODULE_ROOT`:

```powershell
& ./Build/Studio/Apple/Invoke-PinnedPowerForge.ps1 -Tool Web -ArgumentList @(
    'apple-store-art', 'compose', '--config', 'Build/Studio/Apple/store-art/store-art.json',
    '--out', 'Artifacts/Studio/Apple/screenshots/store-art', '--locales', 'de-DE')
```

`store-art/screenshot-configs.rb` writes one screenshot mapping per platform and locale from the spec, with `maxCount` equal to the device's frame count (never more than 10), and lists them in `ScreenshotConfigPaths`. Run it after changing frames or locales; `--check` fails when they are out of date. `check-appstore-localizations.rb` checks that the listing languages match the String Catalogs and that every listing, caption file and release path is present and within App Store limits. The Studio Apple workflow runs both checks on every pull request.

With English captured, the script also writes website candidates (`website-screenshot-assets.sh`) and a contact sheet (`screenshot-contact-sheet.swift`) under `screenshots/website-candidate`. After reviewing them, `website-screenshot-promote.sh --check --source-commit <sha>` compares the candidates from a master capture against `Website/static/images/studio`, and `--apply` copies them as `apple-<device>-<scene>.png`. Neither step commits or deploys the site.

Scan and conversion need the camera or a picker sheet, so they have no store scene yet.

## Store declarations

`powerforge.store-onboarding.json` owns remote setup without archive targets. `appstoreconnect-governance.json` declares free pricing and availability outside France. The initial release excludes France while the French encryption filing is completed. Its empty `encryptionDeclarations` array contains no encryption-document declarations. Both Info.plists set `ITSAppUsesNonExemptEncryption` to true, because the engine implements PDF password encryption (AES and RC4) rather than only calling the system's; App Store Connect still asks for the matching export-compliance answer until a compliance code is added. Empty accessibility and subscription arrays make no claims and erase nothing. Plan and apply governance through PowerForge, retaining the exact reviewed plan.

The native privacy manifest records required-reason APIs used by document recovery, selected files, preferences and timing. The public policy is maintained in `OfficeIMO.Studio/PRIVACY.md` and [the website privacy page](https://officeimo.com/studio/privacy/). Review the native app's actual behavior when answering Apple's privacy questionnaire.

The bundled engine implements PDF encryption and certificate signatures. MIT/public source does not establish export exemption. Apple's declaration and any required documents must be accepted before the relevant review gate; do not infer readiness from a locally generated manifest.

## TestFlight automation

The **TestFlight** GitHub workflow provides **Build now**, **Automatic: Off**, **Automatic: Selected PRs**, and **Automatic: Every merged PR**. It records a request for the trusted PowerForge worker; GitHub does not receive Apple signing credentials. The worker qualifies the exact `master` commit with **Native Apple qualification** and deploys both native targets to the configured internal group. Automatic builds start disabled. Private worker enrollment and operator paths stay outside this public repository.

The **Studio Apple monitor** runs the shared read-only Doctor daily and retains its receipt. Its incident is updated only when the observed release state changes. The worker and monitor use the same pinned PowerForge source.

## Open-source distribution

The first-party software remains MIT licensed. Store pricing is a separate product choice. Apple distribution does not replace the copyright notices, permission grants or third-party license obligations. See [open-source notices](../../../OfficeIMO.Studio/OPEN_SOURCE.md).
