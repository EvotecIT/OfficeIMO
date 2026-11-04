# HEIC to PDF verification

This opt-in source smoke stays outside the normal solution. It explicitly registers the opt-in ImageIO provider, decodes the basic, rotated and 10-bit P3 fixtures, and creates PDFs through `sources.ToPdf()`. On macOS it also writes independent full-file ImageIO references. It uses no third-party HEVC decoder.

From the repository root, choose a new output directory:

```sh
dotnet run --project Build/Verification/HeicSmoke/HeicSmoke.csproj -c Release -f net10.0 -- OfficeIMO.TestAssets/Images/Heic Artefacts/HeicSmoke
mutool draw -q -r 96 -o Artefacts/HeicSmoke/basic-8bit-srgb.rendered.png Artefacts/HeicSmoke/basic-8bit-srgb.pdf
dotnet run --project Build/Verification/HeicSmoke/HeicSmoke.csproj -c Release -f net10.0 --no-build -- --compare Artefacts/HeicSmoke/basic-8bit-srgb.reference.png Artefacts/HeicSmoke/basic-8bit-srgb.rendered.png
```

Repeat rendering and comparison for `orientation-6` and `display-p3-10bit`. At 96 DPI each PDF page has the source pixel dimensions. The comparison requires equal dimensions, mean channel error at most 2/255, and at most 1% of channels differing by more than 8/255. Native chroma interpolation can differ around rotated edges; this is measured separately from orientation and layout. Inspect the rendered rotation chart: its arrow points up and the letter F reads normally.

The program explicitly registers the opt-in ImageIO provider. Run it on macOS; it neither installs a decoder nor enables a route in an application or command host. Other .NET hosts have no default codec and report `image-codec-unavailable`. PSWriteOffice keeps this codec-dependent route unavailable, including planning and `WhatIf`.

For a macOS NativeAOT check, the selected target is applied to the whole source graph during the opt-in publish:

```sh
dotnet publish Build/Verification/HeicSmoke/HeicSmoke.csproj -c Release -f net10.0 -r osx-arm64 -p:PublishAot=true -p:TargetFrameworks=net10.0 -o Artefacts/HeicSmoke/native
Artefacts/HeicSmoke/native/HeicSmoke OfficeIMO.TestAssets/Images/Heic Artefacts/HeicSmoke/native-output
```

Keep native build output in task-owned scratch when running repeated qualification. A macOS NativeAOT pass does not establish iOS device behavior. Camera-photo pickup tests and their fixture requirements are documented in [the asset directory](../../../OfficeIMO.TestAssets/Images/Heic/README.md).
