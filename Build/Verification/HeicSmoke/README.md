# HEIC to PDF verification

This opt-in source smoke stays outside the normal solution. It registers the system provider, decodes the basic, rotated and 10-bit P3 fixtures, and creates PDFs through `PdfDocument.CreateFromImages`. On macOS it also writes independent full-file ImageIO references. It uses no third-party HEVC decoder.

From the repository root, choose a new output directory:

```sh
dotnet run --project Build/Verification/HeicSmoke/HeicSmoke.csproj -c Release -f net10.0 -- OfficeIMO.TestAssets/Images/Heic Artefacts/HeicSmoke
mutool draw -q -r 96 -o Artefacts/HeicSmoke/basic-8bit-srgb.rendered.png Artefacts/HeicSmoke/basic-8bit-srgb.pdf
dotnet run --project Build/Verification/HeicSmoke/HeicSmoke.csproj -c Release -f net10.0 --no-build -- --compare Artefacts/HeicSmoke/basic-8bit-srgb.reference.png Artefacts/HeicSmoke/basic-8bit-srgb.rendered.png
```

Repeat rendering and comparison for `orientation-6` and `display-p3-10bit`. At 96 DPI each PDF page has the source pixel dimensions. The comparison requires equal dimensions, mean channel error at most 2/255, and at most 1% of channels differing by more than 8/255. Native chroma interpolation can differ around rotated edges; this is measured separately from orientation and layout. Inspect the rendered rotation chart: its arrow points up and the letter F reads normally.

On Windows, first run the [WIC qualification commands](../../../OfficeIMO.Drawing.Wic/README.md#windows-validation). The smoke writes the same PDF and composed PNG artifacts; the WIC test project owns the independent full-file WIC reference. macOS references are not used as a Windows oracle.

The PowerShell front door is tested through the built module:

```powershell
./Build/Verification/Test-HeicConversion.ps1 -ModulePath ./PowerShell/PSWriteOffice/PSWriteOffice.psd1 -InputPath ./OfficeIMO.TestAssets/Images/Heic/basic-8bit-srgb.heic -OutputPath ./Artefacts/HeicSmoke/powershell.pdf
```

The command on this branch is `ConvertTo-OfficeDocument`. The script checks its successful workflow result and PDF artifact. Requires PowerShell 7.4 or newer and the installed OS codec; it does not overwrite existing output.

For a macOS NativeAOT check, the selected target is applied to the whole source graph during the opt-in publish:

```sh
dotnet publish Build/Verification/HeicSmoke/HeicSmoke.csproj -c Release -f net10.0 -r osx-arm64 -p:PublishAot=true -p:TargetFrameworks=net10.0 -o Artefacts/HeicSmoke/native
Artefacts/HeicSmoke/native/HeicSmoke OfficeIMO.TestAssets/Images/Heic Artefacts/HeicSmoke/native-output
```

Keep native build output in task-owned scratch when running repeated qualification. A macOS NativeAOT pass does not establish iOS device behavior. Camera-photo pickup tests and their fixture requirements are documented in [the asset directory](../../../OfficeIMO.TestAssets/Images/Heic/README.md).
