# Blazor WebAssembly Conversion Proof

Date: 2026-06-22

This note records a local proof that OfficeIMO document-to-PDF conversion can run from a static Blazor WebAssembly app, which is the deployment model used by GitHub Pages.

## Scope

The proof used a standalone `net10.0` Blazor WebAssembly app that referenced the local OfficeIMO projects directly:

- `OfficeIMO.Word.Pdf`
- `OfficeIMO.Excel.Pdf`
- `OfficeIMO.PowerPoint.Pdf`

The app loaded Office fixtures as browser static assets, opened them from byte arrays/streams, converted them to PDF bytes in the browser runtime, and checked that the output started with `%PDF`.

## Commands

```powershell
dotnet new blazorwasm -o Proofs\BlazorWasmConversionProof -f net10.0
dotnet build Proofs\BlazorWasmConversionProof\BlazorWasmConversionProof.csproj -c Release
dotnet publish Proofs\BlazorWasmConversionProof\BlazorWasmConversionProof.csproj -c Release -o artifacts\blazor-wasm-conversion-proof
dotnet run --project Proofs\BlazorWasmConversionProof\BlazorWasmConversionProof.csproj --configuration Release --urls http://127.0.0.1:5179
```

Browser execution was verified with Playwright against `http://127.0.0.1:5179`.

## Result

| Conversion | Browser result | Notes |
| --- | --- | --- |
| Word DOCX to PDF, basic fixture | Pass | Produced 5,256 PDF bytes after fixing a null settings-part read in `WordSection.DifferentOddAndEvenPages`. |
| Word DOCX to PDF, empty fixture | Pass | Produced 758 PDF bytes. |
| Word DOCX to PDF, sample1 fixture | Expected fail | PDF text encoding preflight rejected U+F0B7 because embedded Unicode fonts are required. This is a font/Unicode coverage gap, not a browser-platform failure. |
| Excel XLSX to PDF | Pass | Produced 7,257 PDF bytes. |
| PowerPoint PPTX to PDF | Pass | Produced 2,088 PDF bytes. |

## Conclusion

GitHub Pages can plausibly host a Blazor WebAssembly conversion playground for OfficeIMO conversions that stay inside byte/stream APIs and do not depend on server processes, native graphics libraries, Office, LibreOffice, or filesystem-only workflows.

The current proof says Excel-to-PDF and PowerPoint-to-PDF already work in-browser for representative fixtures. Word-to-PDF works for simple fixtures after the null guard, but richer Word documents still need the Unicode font embedding path to support symbols such as private-use bullet glyphs.

Before presenting this as a public production feature, validate:

- browser bundle size and startup time;
- memory use for larger DOCX/XLSX/PPTX files;
- explicit browser-safe conversion API wrappers;
- Unicode font embedding or browser-provided font packaging for Word output;
- static-site UX for drag/drop, per-format diagnostics, and download URLs.
