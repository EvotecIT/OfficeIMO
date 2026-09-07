# Labelled scan quality fixtures

The TIFF scans and transcripts come from the pinned Tesseract test repository revision in `sources.json`, under the included Apache-2.0 license. `phototest` contains English text. `eurotext` contains Latin-script text in several European languages. Their original image and transcript SHA-256 hashes are retained in the manifest.

`generate.py` uses Pillow to create PNG variants with a clockwise 3-degree skew, a horizontal paper shadow, a 90-degree turn, or a 180-degree turn. It also writes image-only PDFs at 300 DPI without using OfficeIMO. The original eurotext scan already has a small natural skew, so its final line angle includes that skew and any added rotation.

Run the opt-in integration lane from the repository root with Tesseract and its `eng` and `osd` models installed:

```powershell
dotnet run --project Build/PdfQualityCorpus/OfficeIMO.PdfQualityCorpus.Tool.csproj -c Release -f net8.0 -- scan OfficeIMO.TestAssets/ScanQuality artifacts/scan-quality
```

The runner processes ten fixtures, before and after cleanup, and writes the OCR rasters, searchable PDFs, provider text, reconstructed text, and `scan-quality.json`. The report includes source/truth hashes, orientation evidence, transformation reports, and managed-buffer estimates. Character error rate uses Unicode code points after NFC and whitespace normalization; word error rate uses whitespace-delimited tokens. Both retain case and punctuation, and may exceed 100% when insertions exceed the reference length.

The run fails if the searchable layer changes the original page appearance or source snapshot. Accuracy is reported separately: a low-confidence orientation can intentionally remain unchanged, and an English model does not establish multilingual model quality. These two source scans are regression evidence, not a general OCR accuracy estimate. The normal unit suite uses the checked-in fixtures without launching Tesseract or Python.
