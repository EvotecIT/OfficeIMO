# OfficeIMO Studio native Apple apps design

OfficeIMO Studio on macOS, iPadOS and iOS is a set of native apps written in Swift, using SwiftUI and AppKit or UIKit where a platform needs them. Each app uses the OfficeIMO engines and Studio's document behaviour through one C# library compiled with NativeAOT. Windows and Linux keep the Avalonia Studio. This document describes the boundary between the two halves. The library's current contract is the header [`officeimo_studio.h`](../OfficeIMO.Studio.Native/include/officeimo_studio.h), and open work lives in the [Studio roadmap](ROADMAP.md#phase-7-native-apple-apps).

## Decision

Native apps get each platform's conventions without imitating them: document browsers and `NSDocument`/`UIDocument` lifecycles, the menu bar, Files and the share sheet, security-scoped bookmarks, Pencil, pointer and keyboard on iPad, size classes from iPhone to a large display, VoiceOver, and system typography. The document engines stay in C#, because rewriting them would duplicate the formats, fidelity rules, preservation reports and tests that already exist.

Three alternatives were considered:

| Alternative | Why it was not chosen |
| --- | --- |
| Keep Avalonia on Apple platforms | One codebase, but the result follows Avalonia's controls rather than the platform's. Native menus, the document lifecycle, sandbox storage and iPhone layouts all need per-platform work that a native app gets from the system |
| Native UI written in C# through .NET's iOS and macOS bindings | Viable, and the C interface below keeps it possible. Swift was chosen for SwiftUI, previews, Xcode tooling and the platform's own documentation and samples |
| Rewrite the engines in Swift | Duplicates every format engine and its qualification, and splits behaviour between platforms |

The Avalonia macOS build and the Avalonia iOS host remain in the repository until the native apps cover the everyday journeys. They receive fixes, not new Apple-specific features.

## Layers

| Layer | Owner | Notes |
| --- | --- | --- |
| Presentation, navigation, menus, gestures, accessibility, localization | Swift app per platform | Platform idioms; strings in the app's String Catalog |
| File access: document browser, security-scoped URLs, bookmarks, recent documents, file coordination, iCloud | Swift app | The engine only receives paths the app can currently read or write |
| C interface: handles, memory, status codes, cancellation, JSON payloads | `OfficeIMO.Studio.Native` | Thin; no document logic |
| Document sessions, editing, undo/redo, recovery copies, verified saving, workflows | `OfficeIMO.Studio.Core`, `OfficeIMO.Workflows` | The same types the desktop Studio uses |
| Formats, rendering, conversion, OCR geometry | OfficeIMO engines | Unchanged |

Document behaviour belongs in Studio.Core, and the C interface exposes it. If a native app needs behaviour the desktop already has, such as an identity check before saving, the interface exposes the existing implementation rather than re-implementing it in Swift.

## C interface rules

- **Calls block.** The app calls from a background task, never the main actor. The engine serializes calls on one document, and separate documents run concurrently. A Swift `actor` per document is the natural wrapper.
- **Status and message.** Fallible calls return an `officeimo_status` and, on failure, a UTF-8 message the app may show. Exceptions never cross the interface. Password prompts, unsupported content and cancellation have their own statuses so the app can respond without parsing text.
- **Memory.** Strings, JSON and image bytes returned by the library are released with `officeimo_free`. Each handle (document, cancellation) has its own release function and is released exactly once. Inputs are borrowed for the duration of the call.
- **Structured results** are JSON produced by source-generated serializers, so NativeAOT needs no reflection. Field names are camelCase and stable within an ABI version.
- **Versioning.** Adding a function or an optional JSON field is compatible. Changing a signature, removing a field or changing its meaning raises `officeimo_abi_version`, and the app checks it at launch against `OFFICEIMO_ABI_VERSION`.
- **Cancellation** tokens may be cancelled from any thread while a call runs.
- **Initialization.** `officeimo_initialize` names an app-owned folder inside the container for recovery copies. Opening a document before that fails with `OFFICEIMO_NOT_INITIALIZED`.

## Storage boundary

The app owns every grant to user files. To open a document it resolves the URL, which might come from the document browser, a bookmark, Files or a share. It then starts security-scoped access, passes the path, and keeps access open for the call. The engine never retains access, bookmarks or recent-document lists.

Saving writes beside the destination and replaces it only once the output is complete, so a failed save leaves the previous file. An in-place save refuses when the file was replaced or moved since it was opened, which can happen with iCloud or another app. The app then offers to save elsewhere or reload. The app wraps saves in its file-coordination and document lifecycle: a coordinated write for replacing, version conflicts and autosave policy are host decisions.

Recovery copies of unsaved edits live under the folder given to `officeimo_initialize`. The document facts report `hasRecovery` when an earlier session left edits for the opened file.

## Rendering

A page renders as PNG at the scale the app requests, in pixels per PDF point. The app supplies its display scale and decodes the PNG with Image I/O. This keeps the first interface simple. Interactive zoom, tiling and avoiding the encode/decode step are measured in the feasibility work before the interface gains raw pixel buffers or tiles.

## Building the engine

On a desktop, publish a shared library and run the [document journey script](../OfficeIMO.Studio.Native/Test-NativeEngine.ps1) against it:

```sh
dotnet publish OfficeIMO.Studio.Native/OfficeIMO.Studio.Native.csproj -c Release -r osx-arm64 -o artifacts/engine/osx-arm64
pwsh OfficeIMO.Studio.Native/Test-NativeEngine.ps1 -Library artifacts/engine/osx-arm64/OfficeIMOStudioEngine.dylib -Pdf sample.pdf -WorkDirectory /tmp/engine-check
```

For an app, publish static libraries with `-p:NativeLib=Static` for `ios-arm64`, `iossimulator-arm64`, `osx-arm64` and `osx-x64`. Then combine them with the header into an XCFramework using `xcodebuild -create-xcframework`. This requires a Mac with Xcode. A static NativeAOT library also needs the runtime's own native libraries and system frameworks at link time. The feasibility work records the exact link inputs and publishes them with the framework.

The library references Studio.Core and the engines below it; it contains no Avalonia or other UI toolkit. No entry point launches a process: the desktop-only process adapters in that graph, such as CUPS printing, stay unreachable through the interface, matching the sandboxed and mobile policy.
