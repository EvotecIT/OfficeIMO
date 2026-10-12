# OfficeIMO Studio for macOS and iPadOS

The native Swift app for macOS and iPadOS, built on the shared C# engine through its C interface. It currently runs the engine journey: it starts the engine with a folder in the app container, opens a bundled PDF, shows page 1, rotates it, undoes and redoes the rotation, and saves a copy. The [native Apple apps design](../../Docs/officeimo.studio-apple-native-design.md) describes the boundary between the app and the engine.

## Build

The app links `OfficeIMOStudioEngine.xcframework`, which is built from source and kept out of the repository. On a Mac with Xcode and the .NET SDK selected by `global.json`, from the repository root:

```sh
pwsh OfficeIMO.Studio.Native/Build-AppleEngine.ps1 -OutputDirectory <build-folder>/engine
mkdir -p Apple/OfficeIMOStudio/Engine
ln -sfn <build-folder>/engine/OfficeIMOStudioEngine.xcframework Apple/OfficeIMOStudio/Engine/OfficeIMOStudioEngine.xcframework
```

Then open `OfficeIMOStudio.xcodeproj` and run the `OfficeIMOStudio-macOS` or `OfficeIMOStudio-iOS` scheme. The framework's module map links the system libraries and frameworks the engine needs, so the targets only link the XCFramework.

## Layout

| Folder | Contents |
| --- | --- |
| `Shared/Engine` | `Engine` for process-wide calls and `EngineDocument`, an actor per open document. Engine calls block, so each document actor runs on its own serial queue, off the main actor and Swift's cooperative pool. Task cancellation cancels the engine call. |
| `Shared/Journey` | The journey model and view. |
| `Shared/Measurement` | Process metrics and the feasibility measurements. |
| `Measurements` | The measurement documents' generator and the runner for the Mac and an iPad simulator. |

## Unattended runs

| Launch argument | Effect |
| --- | --- |
| `-OfficeIMOJourney` | Runs the journey and writes `Documents/Results/journey.json`. |
| `-OfficeIMOMeasure <file>` | Measures `Documents/Measurements/<file>` and writes `Documents/Results/measure-<file>.json`. |
| `-OfficeIMOExit` | Quits when done. |

Results are also printed to standard output. `Measurements/Measure-EngineProof.ps1 -WorkDirectory <folder>` builds the measurement PDFs and the app, runs the journey and the measurements on the Mac and in a temporary iPad simulator, and reads the engine's share of the app size from the linker map.
