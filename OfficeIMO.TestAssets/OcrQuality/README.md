# Native OCR quality corpus

This opt-in runner measures real, independently produced inputs with hash-bound gold text. It compares baseline Tesseract recognition with bounded segmentation retries, records geometry and source preservation, and counts confidence checks that pass despite incorrect text. It does not qualify hosted AI answers or general document accuracy.

## Run

Install a native Tesseract executable, then run from the repository root:

```powershell
dotnet run --project Build/PdfQualityCorpus/OfficeIMO.PdfQualityCorpus.Tool.csproj -f net10.0 -- ocr-quality OfficeIMO.TestAssets/OcrQuality/manifest.json OfficeIMO.TestAssets <new-output-directory>
```

The existing language provisioner downloads checksum-pinned `tessdata_fast` models for English, Spanish, German, and French. The scorecard records model digests, catalog revision, native executable version, framework, operating system, and architecture. PDF regions render through OfficeIMO at 300 DPI; the two pinned single-page TIFF scans pass their original bytes directly to the native provider. The shadow/skew PNG uses the existing scan cleanup owner. Each case runs twice over the same input, with page segmentation modes 3, 6, and 11 available under one 45-second adaptive budget.

The default exit code gates completed recognition, source hashes, and word geometry inside the raster. Accuracy limits remain measured failures. Append `--require-quality` to require every declared CER and WER limit as well. Exits are 0 for the selected gate passing, 1 for a failed gate, 2 for setup or manifest errors, and 3 for cancellation. Ctrl+C and the 15-minute run deadline prevent publication of a final scorecard. Partial text and raster evidence can remain for diagnosis. An existing final scorecard is never overwritten.

The [native qualification workflow](../../.github/workflows/ocr-native-quality.yml) runs the shared OCR contracts and this operational gate on hosted Windows, Linux, and macOS. Only a completed platform scorecard establishes execution evidence for that platform; the workflow definition alone does not.

## Inputs and labels

`manifest.json` pins every source and `labels.json` by SHA-256 before recognition. Scoped Git attributes preserve those bytes during Windows checkout. Sources remain unchanged. Labels and accuracy limits are declared independently of provider output:

| Cases | Source and gold | Boundary |
| --- | --- | --- |
| English and Spanish headings, certification, voucher heading, unit-price region, checkbox labels | Original [IRS W-9](https://www.irs.gov/pub/irs-pdf/fw9.pdf), [Spanish W-9](https://www.irs.gov/pub/irs-pdf/fw9sp.pdf), [W-8BEN](https://www.irs.gov/pub/irs-pdf/fw8ben.pdf), and [GSA SF1034](https://www.gsa.gov/system/files/SF1034-87c.pdf); gold transcribed from independently rendered visible regions before OCR | Blank government forms and a blank voucher; no filled invoice or field-value extraction claim |
| Four-column magazine and travel magazine scans | Original TIFFs and complete upstream transcripts from the [pinned Tesseract test repository](https://github.com/tesseract-ocr/test/tree/232ff181c66516116ec0e84c4963f70de15050fd/testing); `8071.txt` and `8087.txt` preserve upstream bytes | Recognition and reading sequence; embedded publication credits retained |
| Photo text, shadow/skew variant, multilingual text | Existing [ScanQuality sources](../ScanQuality/sources.json), prepared PNGs and upstream gold text | Reuses recorded conversion/preparation; no new synthetic gold derived from recognition |

The four PDFs are United States federal government forms. The upstream OCR fixtures retain the test distribution's [Apache 2.0 license](LICENSE.tesseract-test) and original embedded publication credits. Full source URLs and hashes are in the manifest. PDF label regions use normalized top-left page coordinates and deliberately omit nearby fields or borders where stated by the crop. A cropped label does not qualify whole-page layout or reading order. The table-region case measures text only, not recovered table structure or cell geometry.

## Interpret the scorecard

Every case retains baseline and selected text, raster and source hashes, elapsed time, attempted variants, uncertainty counts, disagreement, review recommendation, source preservation, and geometry checks. Repeated-output stability checks normalized selected text and prepared raster hashes. The report also groups failures and changes by document class and compares confidence thresholds 0.7, 0.8, 0.9, and 0.95 against independently declared accuracy limits.

CER uses Unicode code points; WER uses whitespace-separated tokens. Both apply NFC and collapse whitespace while preserving case and punctuation. Word edit alignment reports insertions, deletions, and substitutions, with deterministic tie handling. Deletions can arise from reading-order mistakes and are not a semantic fact-omission metric. Rates may exceed 100% when output adds enough text. Comparisons have an explicit edit-distance work bound and observe cancellation.

A confidence false pass means the word-evidence checks passed while CER or WER exceeded the gold limit. It does not measure whether a human would accept the result. Host peak working set is cumulative across the run and excludes the native child process; it is not per-document or provider peak memory. This small corpus is not a held-out calibration set. Thresholds should be evaluated on the intended workload before adoption.

Remaining product work belongs in the [single roadmap](../../Docs/ROADMAP.md#document-assistant): filled invoices and independent structured labels, broader scripts and layouts, semantic support and omission, hosted model profiles, provider memory, and native Studio journeys.
