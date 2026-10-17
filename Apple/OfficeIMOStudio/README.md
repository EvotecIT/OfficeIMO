# OfficeIMO Studio for macOS, iPadOS and iOS

The native Swift app for the Mac, iPad and iPhone, built on the shared C# engine through its C interface. It opens PDFs as system documents with the Read, Markup and Pages tasks: search, text selection and markup, page organization, form filling, handwritten and certificate signatures, password protection, on-device text recognition, scanning, conversion to PDF, printing, sharing and Shortcuts actions. Pages render as the engine's pixel rectangles in the system's zooming scroll view, and layouts follow the width, so the same scene serves the Mac, iPad, iPhone and a folding iPhone. The [feature table on the Studio page](https://officeimo.com/studio/#features), from [`Website/data/studio_features.json`](../../Website/data/studio_features.json), compares the app with Studio on Windows and Linux. The [native Apple apps design](../../Docs/officeimo.studio-apple-native-design.md) describes the interface and the boundary between the app and the engine.

## Build

Open `OfficeIMOStudio.xcodeproj` and run the `OfficeIMOStudio-macOS` or `OfficeIMOStudio-iOS` scheme. Both targets use the local [engine package](../OfficeIMOStudioEngine/Package.swift), which downloads the versioned OfficeIMO NativeAOT XCFramework and verifies its SHA-256. No separate local engine copy or symlink is needed. The apps require macOS, iPadOS or iOS 26; the framework includes Intel and Apple silicon Mac slices, arm64 iOS, and the arm64 iOS simulator.

The Swift UI calls the existing OfficeIMO engine through its C interface. Packaging the engine separately changes build delivery, not document behavior or the native interface. The package also embeds the engine's MIT and third-party license notices.

To produce an updated engine from the repository root on a Mac with Xcode and the SDK selected by `global.json`:

```sh
pwsh OfficeIMO.Studio.Native/Build-AppleEngine.ps1 -OutputDirectory <build-folder>/engine
pwsh Build/Studio/Apple/Pack-StudioAppleEngine.ps1 -EngineDirectory <build-folder>/engine -OutputDirectory <build-folder>/package -PowerForgeCli <PowerForge.Cli.dll> -SourceCommit <engine-source-commit> -Version <x.y.z>
```

The packer uses PowerForge's exact dependency notice validation and archive builder. Its `Publish` option uses PowerForge's GitHub release publisher and the private `GITHUB_TOKEN` environment variable. Update the package's immutable URL, checksum and generated license resources together for a new engine artifact. See [Apple distribution](../../Build/Studio/Apple/README.md) for release entrypoints.

## Layout

The `Shared` folder is synchronized with both targets, so new files need no project edits.

| Folder | Contents |
| --- | --- |
| `Shared/App` | The app's scenes and menu commands, unattended runs and the development harness. |
| `Shared/Document` | `StudioDocument`, the system document over a private working copy that the engine edits, and `DocumentSession`, a window's state: the open engine document, facts, task, page, zoom, selection and undo registration. |
| `Shared/Canvas` | The page view: the page layout in PDF points, the tile renderer, and its hosts in `UIScrollView` and `NSScrollView`. |
| `Shared/Views` | The document window, toolbar, page list, Pages task and inspector. |
| `Shared/Conversion` | Converting Word, Excel, PowerPoint and other documents to PDF with the engine, from the launch screen, the File menu and Shortcuts. |
| `Shared/Forms` | Filling form fields in the Read task: the engine's fields per page, the text, check box, radio and choice editors over the page in UIKit and AppKit, and each value as one undoable change. |
| `Shared/Signatures` | Handwritten signatures kept on the device, the pad for drawing one, and placing one on a page as ink. |
| `Shared/Intents` | Shortcuts actions and their Siri and Spotlight phrases, over `DocumentActions`, which the unattended journey also runs. |
| `Shared/Printing` | Printing the engine's page renderings through the system print panel and sheet. |
| `Shared/Scan` | Scanning paper into a searchable PDF: the VisionKit document camera and the photo picker on iPad and iPhone, image files and Continuity Camera on the Mac, and Vision text recognition passed to the engine. |
| `Shared/Engine` | `Engine` for process-wide calls and `EngineDocument`, an actor per open document. Engine calls block, so each document actor runs on its own serial queue, off the main actor and Swift's cooperative pool. Task cancellation cancels the engine call. Page rectangles become `CGImage`s that own the engine's pixels. |
| `Shared/Journey`, `Shared/Measurement` | The engine journey and the measurements for unattended runs. |
| `Measurements` | The measurement documents' generator and the runner for the Mac and an iPad simulator. |

`Shared/Localizable.xcstrings`, `InfoPlist.xcstrings` and `AppShortcuts.xcstrings` hold the app's strings in English, Polish, German, French, Spanish, Italian and Brazilian Portuguese for both targets, reusing the wording of OfficeIMO Studio's [own translations](../../OfficeIMO.Studio/Localization) where the two apps share a term. Xcode adds new strings to `Localizable.xcstrings` when it builds. From the command line, build both targets and sync the catalog with the extracted strings:

```sh
xcrun xcstringstool sync Shared/Localizable.xcstrings --stringsdata <derived data>/Build/Intermediates.noindex/OfficeIMOStudio.build/*/*.build/Objects-normal/arm64/*.stringsdata
```

## State contract tests

The Foundation state contracts can also be tested without building the app or the native engine. From the repository root on a Mac:

```sh
xcrun swift test --package-path Apple --scratch-path <test-output>/swift-state
```

These tests cover export permission selection, revision-bound reads and redaction search state. They do not qualify SwiftUI interactions or the native apps. To produce an XCTest report with Swift 6.4, add `--parallel --num-workers 1 --disable-swift-testing --xunit-output <test-output>/swift-state.xml`.

## Unattended runs

| Launch argument | Effect |
| --- | --- |
| `-OfficeIMOJourney` | Runs the journey and writes `Documents/Results/journey.json`. |
| `-OfficeIMOMeasure <file>` | Measures `Documents/Measurements/<file>` and writes `Documents/Results/measure-<file>.json`. |
| `-OfficeIMOExit` | Quits when done. |

Results are also printed to standard output. `Measurements/Measure-EngineProof.ps1 -WorkDirectory <folder>` builds the measurement PDFs and the app, runs the journey and the measurements on the Mac and in a temporary iPad simulator, and reads the engine's share of the app size from the linker map.

## Development harness

Debug builds for iPad and iPhone accept `-OfficeIMOOpen <file>`, which opens `Documents/<file>` from the app's container straight into a document window, without the system's document browser. Use it when the simulator's document browser cannot reach files, as happens when the simulator's data lives on an external volume:

```sh
xcrun simctl launch booted com.evotec.officeimo.studio -OfficeIMOOpen sample.pdf
```

Copy the PDF into the container first; `xcrun simctl get_app_container booted com.evotec.officeimo.studio data` prints its folder. The harness shows the same window as a system document but saves nothing back.

`-OfficeIMOScanFromPhotos` shows the launch screen's **PDF from Photos** action on its own and opens the scan it makes in the same kind of window. Add a photo of a printed page with `xcrun simctl addmedia booted <image>` first; the simulator has no document camera.
