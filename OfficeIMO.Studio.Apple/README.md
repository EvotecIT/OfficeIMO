# OfficeIMO Studio for Apple platforms

This native prototype opens PDFs, displays their pages, adds text notes through OfficeIMO.Pdf, and saves the resulting document. It targets Apple silicon macOS 26 and iOS/iPadOS 26. The SwiftUI interface uses system navigation, toolbars, sheets, document pickers, and glass page controls. PDFKit provides the display, selection, scrolling, and zoom.

Use **New Document** to try the bundled PDF or open a local PDF. **Add Note** places a note near the lower-left corner of the selected page. Read notes in the page sidebar. SwiftUI manages document saving and unsaved-window handling; **Save a Copy** exports a separate PDF. On compact displays, Pages and Notes opens in a sheet so the document keeps the full screen width.

## Scope

- The C# engines remain shared with the other OfficeIMO surfaces. `OfficeIMO.Studio.Native` exposes a small, stateless C ABI compiled with .NET NativeAOT.
- PDFKit is a read-only display model. OfficeIMO owns annotation changes and the saved bytes; SwiftUI owns file coordination and undo registration.
- Input and output are limited to 64 MiB. Notes accept up to 4,000 UTF-16 code units. Password-protected PDFs are not supported.
- This slice supports reading and adding text notes. It does not implement the full desktop editor, form filling, Pencil ink, OCR, conversion, or batch jobs.
- The prototype uses `com.evotec.officeimo.studio.nativeprototype`. The production Store identity and Avalonia application remain in their existing projects.

The [Studio roadmap](../Docs/ROADMAP.md#phase-7-ipad-then-iphone) owns the remaining native migration, device qualification, and distribution work. A native framework build is not App Store qualification. Physical iPad journeys, file-provider interruption/recovery, accessibility, signed distribution, and exact-binary notices remain release gates.

## Build the prototype

Use an Apple silicon Mac, Xcode 26 or newer, XcodeGen, and the .NET SDK selected by the repository's `global.json`. Run from the repository root. Choose a fresh output directory on the configured development volume:

```sh
export STUDIO_BUILD_ROOT="${EVOTEC_DEV_TEMP:?Set the development scratch root}/officeimo-native-build"
export OFFICEIMO_NATIVE_FRAMEWORK_ROOT="$STUDIO_BUILD_ROOT/frameworks"
```

Publish the C# library for each required platform. These are native libraries; they do not require a separate .NET iOS app host or iOS workload.

```sh
dotnet publish OfficeIMO.Studio.Native -c Release -r osx-arm64 --artifacts-path "$STUDIO_BUILD_ROOT/dotnet" -o "$STUDIO_BUILD_ROOT/engine/mac"
dotnet publish OfficeIMO.Studio.Native -c Release -r iossimulator-arm64 --artifacts-path "$STUDIO_BUILD_ROOT/dotnet" -o "$STUDIO_BUILD_ROOT/engine/simulator"
dotnet publish OfficeIMO.Studio.Native -c Release -r ios-arm64 --artifacts-path "$STUDIO_BUILD_ROOT/dotnet" -o "$STUDIO_BUILD_ROOT/engine/device"
```

For this prototype, assemble the frameworks using the standard layout from Microsoft's [NativeAOT framework guide](https://learn.microsoft.com/en-us/dotnet/core/deploying/native-aot/ios-like-platforms/creating-and-consuming-custom-frameworks). The checked-in plists declare the platform and minimum OS. Production framework assembly and signing belong in the shared PowerForge Apple pipeline.

Mac framework:

```sh
mkdir -p "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/mac/OfficeIMOStudioEngine.framework/Versions/A/Resources"
cp "$STUDIO_BUILD_ROOT/engine/mac/OfficeIMOStudioEngine.dylib" "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/mac/OfficeIMOStudioEngine.framework/Versions/A/OfficeIMOStudioEngine"
cp OfficeIMO.Studio.Apple/Bridge/MacOSX.plist "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/mac/OfficeIMOStudioEngine.framework/Versions/A/Resources/Info.plist"
ln -s A "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/mac/OfficeIMOStudioEngine.framework/Versions/Current"
ln -s Versions/Current/OfficeIMOStudioEngine "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/mac/OfficeIMOStudioEngine.framework/OfficeIMOStudioEngine"
ln -s Versions/Current/Resources "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/mac/OfficeIMOStudioEngine.framework/Resources"
install_name_tool -id @rpath/OfficeIMOStudioEngine.framework/Versions/A/OfficeIMOStudioEngine "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/mac/OfficeIMOStudioEngine.framework/Versions/A/OfficeIMOStudioEngine"
codesign --force --sign - "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/mac/OfficeIMOStudioEngine.framework"
```

Simulator and device frameworks:

```sh
mkdir -p "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/simulator/OfficeIMOStudioEngine.framework" "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/device/OfficeIMOStudioEngine.framework"
cp "$STUDIO_BUILD_ROOT/engine/simulator/OfficeIMOStudioEngine.dylib" "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/simulator/OfficeIMOStudioEngine.framework/OfficeIMOStudioEngine"
cp "$STUDIO_BUILD_ROOT/engine/device/OfficeIMOStudioEngine.dylib" "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/device/OfficeIMOStudioEngine.framework/OfficeIMOStudioEngine"
cp OfficeIMO.Studio.Apple/Bridge/iPhoneSimulator.plist "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/simulator/OfficeIMOStudioEngine.framework/Info.plist"
cp OfficeIMO.Studio.Apple/Bridge/iPhoneOS.plist "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/device/OfficeIMOStudioEngine.framework/Info.plist"
install_name_tool -id @rpath/OfficeIMOStudioEngine.framework/OfficeIMOStudioEngine "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/simulator/OfficeIMOStudioEngine.framework/OfficeIMOStudioEngine"
install_name_tool -id @rpath/OfficeIMOStudioEngine.framework/OfficeIMOStudioEngine "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/device/OfficeIMOStudioEngine.framework/OfficeIMOStudioEngine"
codesign --force --sign - "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/simulator/OfficeIMOStudioEngine.framework"
codesign --force --sign - "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/device/OfficeIMOStudioEngine.framework"
xcodebuild -create-xcframework -framework "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/simulator/OfficeIMOStudioEngine.framework" -framework "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/device/OfficeIMOStudioEngine.framework" -output "$OFFICEIMO_NATIVE_FRAMEWORK_ROOT/mobile/OfficeIMOStudioEngine.xcframework"
```

Generate the Xcode project and run the Mac contract tests:

```sh
xcodegen generate --spec OfficeIMO.Studio.Apple/project.yml
xcodebuild -project OfficeIMO.Studio.Apple/OfficeIMOStudioApple.xcodeproj -scheme OfficeIMOStudioMac -destination 'platform=macOS,arch=arm64' -derivedDataPath "$STUDIO_BUILD_ROOT/xcode" CODE_SIGN_IDENTITY=- test
```

For the mobile tests, select an installed simulator from `xcrun simctl list devices available` and run the `OfficeIMOStudioMobile` scheme with `-destination 'platform=iOS Simulator,id=<device-UUID>'`. Device builds require development signing and provisioning for the prototype bundle identifier. Ad-hoc Mac and simulator signatures above are local validation signatures, not distribution signatures.

`project.yml` is the maintained project source; XcodeGen output is ignored. The XCTest suite exercises the actual native library, PDFKit readback, invalid inputs, and exact-byte document undo/redo. The bundled `Welcome.pdf` is generated by operation 2 in `StudioNativeExports`, so its document content has a source owner.
