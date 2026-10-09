# Reviewed form OCR consumer

This opt-in console consumer exercises the public `OfficeIMO.Pdf.Ocr` review API with an installed Tesseract provider. It stays outside the default solution and package graph.

Run it with a source PDF, a new task-owned output directory, a JSON file of explicitly reviewed field values, and the installed Tesseract executable path:

```sh
dotnet run --project Build/VerifyPdfFormOcr -- source.pdf review-output accepted-values.json /path/to/tesseract
```

Example decision file:

```json
{"FullName":"Alex Morgan","SerialCode":"A1B2C3","Country":"Poland","Amount":"123.45"}
```

The JSON file supplies acceptance; the consumer never automatically accepts OCR results. It writes the reviewed copy and a compact report with original provider evidence and saved field values. Inspect that output with an independent form reader and renderer to qualify values and appearances. Synthetic producer fixtures and one installed provider run establish a bounded workflow, not general recognition accuracy.
