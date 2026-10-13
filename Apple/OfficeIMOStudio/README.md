# OfficeIMO Studio for macOS, iPadOS and iOS

The native Swift app for the Mac, iPad and iPhone, built on the shared C# engine through its C interface. It opens PDFs as system documents and offers the Read and Pages tasks: pages render as the engine's pixel rectangles in the system's zooming scroll view, with a page list, an inspector of document facts, rotation with system undo, and sharing. Layouts follow the width, so the same scene serves the Mac, iPad, iPhone and a folding iPhone. The [native Apple apps design](../../Docs/officeimo.studio-apple-native-design.md) describes the interface and the boundary between the app and the engine.

## Build

The app links `OfficeIMOStudioEngine.xcframework`, which is built from source and kept out of the repository. On a Mac with Xcode and the .NET SDK selected by `global.json`, from the repository root:

```sh
pwsh OfficeIMO.Studio.Native/Build-AppleEngine.ps1 -OutputDirectory <build-folder>/engine
mkdir -p Apple/OfficeIMOStudio/Engine
ln -sfn <build-folder>/engine/OfficeIMOStudioEngine.xcframework Apple/OfficeIMOStudio/Engine/OfficeIMOStudioEngine.xcframework
```

Then open `OfficeIMOStudio.xcodeproj` and run the `OfficeIMOStudio-macOS` or `OfficeIMOStudio-iOS` scheme. The apps require macOS, iPadOS or iOS 26 and run in the arm64 iOS simulator. The framework's module map links the system libraries and frameworks the engine needs, so the targets only link the XCFramework.

## Layout

The `Shared` folder is synchronized with both targets, so new files need no project edits.

| Folder | Contents |
| --- | --- |
| `Shared/App` | The app's scenes and menu commands, unattended runs and the development harness. |
| `Shared/Document` | `StudioDocument`, the system document over a private working copy that the engine edits, and `DocumentSession`, a window's state: the open engine document, facts, task, page, zoom, selection and undo registration. |
| `Shared/Canvas` | The page view: the page layout in PDF points, the tile renderer, and its hosts in `UIScrollView` and `NSScrollView`. |
| `Shared/Views` | The document window, toolbar, page list, Pages task and inspector. |
| `Shared/Conversion` | Converting Word, Excel, PowerPoint and other documents to PDF with the engine, from the launch screen, the File menu and Shortcuts. |
| `Shared/Intents` | Shortcuts actions and their Siri and Spotlight phrases, over `DocumentActions`, which the unattended journey also runs. |
| `Shared/Printing` | Printing the engine's page renderings through the system print panel and sheet. |
| `Shared/Scan` | Scanning paper into a searchable PDF: the VisionKit document camera and the photo picker on iPad and iPhone, image files and Continuity Camera on the Mac, and Vision text recognition passed to the engine. |
| `Shared/Engine` | `Engine` for process-wide calls and `EngineDocument`, an actor per open document. Engine calls block, so each document actor runs on its own serial queue, off the main actor and Swift's cooperative pool. Task cancellation cancels the engine call. Page rectangles become `CGImage`s that own the engine's pixels. |
| `Shared/Journey`, `Shared/Measurement` | The engine journey and the measurements for unattended runs. |
| `Measurements` | The measurement documents' generator and the runner for the Mac and an iPad simulator. |

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
