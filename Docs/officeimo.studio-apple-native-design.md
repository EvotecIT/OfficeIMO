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

A page renders as PNG at the scale the app requests, in pixels per PDF point. The document facts report `maxRenderPixels`, the largest page image the engine produces (8.4 million pixels). The app clamps its scale to that budget, rounding each side up to whole pixels as the engine does, and a larger request fails with `OFFICEIMO_INVALID_ARGUMENT` before any page content is read. The app decodes the PNG with Image I/O.

The renderer has a separate working budget of 64 million pixels for the page image, intermediate surfaces and decoded source images. Quartz, and so Preview, PDFKit and the Apple document scanner, marks drawn images `/Interpolate true`. The renderer's interpolated image path then needs several times the image's pixels: at a 13-inch iPad's width, a 300 dpi colour letter scan needs 41 million working pixels and an A3 scan 60 million.

### Measurements

[`Measure-EngineProof.ps1`](../Apple/OfficeIMOStudio/Measurements/Measure-EngineProof.ps1) builds the app and runs the journey and the measurements on the Mac and in a temporary iPad simulator. Its measurement documents come from repository fixtures: a 600-page PDF of text, table, form and presentation pages from several producers, and a 40-page scan of 300 dpi colour JPEG pages, each with its own image. Each document is measured in its own launch, three times. The table gives medians from the [run on 2026-10-12](../Apple/OfficeIMOStudio/Measurements/Evidence/2026-10-12/summary.json), Mac mini M4 (`Mac16,10`) with Xcode 27.0. The Mac renders pages 900 pixels wide on its 1x display. The simulator is an iPad Pro 13-inch (M5) on iOS 27.0, rendering pages 2064 pixels wide, its portrait width at 2x.

The host was heavily loaded by other work during the run (load average 360 to 600 on 10 cores), so the timings are upper bounds. The simulator runs on the Mac's CPU and is not device performance. Pixel sizes, budgets and sizes do not depend on load.

| | Mac, long PDF | Mac, scan | Simulator, long PDF | Simulator, scan |
| --- | --- | --- | --- | --- |
| Open | 96 ms | 46 ms | 249 ms | 92 ms |
| First page, open to decoded image | 355 ms | 1.07 s | 1.13 s | 2.99 s |
| Render, median and 90th percentile of 24 pages | 86 / 244 ms | 914 ms / 1.03 s | 620 ms / 1.19 s | 2.32 / 2.85 s |
| PNG decode, median | 4 ms | 4 ms | 20 ms | 27 ms |
| Peak footprint | 172 MB | 747 MB | 270 MB | 625 MB |

The first engine call, which starts the runtime, takes 5 to 15 ms, and 350 ms in the simulator's first launch after it boots. Initialization takes under 40 ms. From process start the app's initializer runs after 35 to 45 ms on the Mac and 250 ms in the simulator. In a release build for iOS devices the engine contributes 26.7 MB of the 27.5 MB the linker places. The stripped executable is 22.7 MB, 9.7 MB compressed.

Splitting one page's render in a NativeAOT build on the Mac, at about 1800 pixels wide, shows where the time goes. For long-PDF pages, reading the page into a drawing takes 5 to 70 ms and rasterizing 68 to 178 ms. Encoding the PNG at the compression the engine uses takes 141 to 261 ms, about half of the render, or 22 to 44 ms at the fastest compression. A raw page buffer is 16.8 MB against 150 to 390 KB of PNG. For scanned pages, decoding the 300 dpi JPEG takes 1.1 to 1.3 s on every render, rasterizing 0.3 to 0.6 s and encoding 0.2 to 0.4 s.

### Decisions

- **Raw pixel buffers for display.** PNG encoding is about half of a text page's render and the app decodes it again. The interface gains a render call that returns premultiplied pixels with their row stride, which the app wraps in a `CGImage` released through `officeimo_free`. PNG stays for thumbnails, sharing and export.
- **Tiles for zoom.** A letter page filling a 13-inch iPad's width is 5.5 million pixels, so zooming past 1.5 times exceeds the page budget. The interface gains region rendering: a rectangle of the page at a scale, each tile within the budget.
- **Reuse the page between renders.** Each render reads the page and decodes its images again, which costs scanned pages over a second before tiles multiply the renders. The document session keeps the page's drawing and decoded images for its revision. Source images decode at the resolution the render needs.

## Building the engine

On a desktop, publish a shared library and run the [document journey](../OfficeIMO.Studio.Native/Test-NativeEngine.ps1) against it:

```sh
dotnet publish OfficeIMO.Studio.Native/OfficeIMO.Studio.Native.csproj -c Release -r osx-arm64 -o artifacts/engine/osx-arm64
pwsh OfficeIMO.Studio.Native/Test-NativeEngine.ps1 -Library artifacts/engine/osx-arm64/OfficeIMOStudioEngine.dylib -Pdf sample.pdf -WorkDirectory /tmp/engine-check
```

The journey is a C program, [`officeimo_journey.c`](../OfficeIMO.Studio.Native/journey/officeimo_journey.c), compiled against the header. It loads the library at run time with each entry point typed from the header, the way a native app hosts it. A .NET process cannot host the library on macOS or Linux: both runtimes handle the same thread-suspension signal, and the first render crashes.

For the apps, [`Build-AppleEngine.ps1`](../OfficeIMO.Studio.Native/Build-AppleEngine.ps1) builds `OfficeIMOStudioEngine.xcframework` on a Mac with Xcode:

```sh
pwsh OfficeIMO.Studio.Native/Build-AppleEngine.ps1 -OutputDirectory <build-folder>/engine
```

It publishes static libraries with `-p:StudioEngineLibrary=Static` for `osx-arm64`, `osx-x64`, `ios-arm64` and `iossimulator-arm64`, then packages them as follows:

- The project sets `NativeLib` itself. As a command-line property it would reach the referenced projects, and on iOS runtimes the Mono workload then treats each of them as a Mono library-mode build. iOS runtimes use the NativeAOT runtime packs, so no workload is needed.
- NativeAOT's static archive holds only the compiled code. A static publish also copies the toolchain's link inputs and lists them in `link/link-inputs.txt`. The script keeps each runtime's list in `link-inputs/<runtime>.txt`.
- The runtime bootstrapper exports nothing and starts the runtime from a static initializer, so as an archive member the linker would drop it. The script prelinks it with `ld -r` into the object that holds the exports, keeping private external symbols visible to the runtime archives. It also marks that object for the apps' deployment targets, macOS 14 and iOS 17.
- Each runtime becomes one static library, merged from that object, the runtime (`libRuntime.WorkstationGC`, `libeventpipe-disabled`, `libstandalonegc-disabled`, `libaotminipal`), the native shims (`libSystem.Native`, `libSystem.Globalization.Native`, `libSystem.IO.Compression.Native`, `libSystem.Net.Security.Native`, `libSystem.Security.Cryptography.Native.Apple`) and Brotli. The two macOS architectures are joined with `lipo`.
- The system inputs go into each slice's module map as link directives. Every slice links `objc`, `swiftCore`, `swiftFoundation`, `icucore`, `dl` and `m`, and the CoreFoundation, CryptoKit, Foundation, Network, Security and GSS frameworks. iOS slices also link `c++` and the system `z`; NativeAOT names the C++ library `stdc++`, which Apple SDKs ship as `libc++`. A Swift target that imports `OfficeIMOStudioEngine` links all of them with no further settings.

The library builds with no trim or AOT warnings for all four runtimes. The slices are 110 MB for iOS and the iOS simulator and 223 MB for universal macOS, including debug information, before the app's linker strips them.

The library references Studio.Core and the engines below it; it contains no Avalonia or other UI toolkit. No entry point launches a process: the desktop-only process adapters in that graph, such as CUPS printing, stay unreachable through the interface, matching the sandboxed and mobile policy.
