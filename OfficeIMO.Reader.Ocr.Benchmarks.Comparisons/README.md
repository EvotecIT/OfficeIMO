# OCR engine comparison

This opt-in suite compares free local OCR engines without adding them to the normal OfficeIMO restore graph. It uses generated, validated English, Polish, dense-table, and low-contrast/skew fixtures. Each lane recognizes the same PNG through a one-shot WSL process boundary, which matches the isolation model available to OfficeIMO providers.

## Support decision

OfficeIMO supports Tesseract as the default local engine. The reusable `IOfficeOcrEngine`, `OfficeOcrEnginePdfProvider`, and `IPdfOcrProvider` seams remain the supported route for caller-owned neural, native, or cloud engines.

RapidOCR is useful and had materially better Polish-diacritic accuracy in this small corpus. It is not part of the default package graph because the tested Python and ONNX deployment was about 425 MiB, took roughly 6 seconds for every one-shot request, and returned line regions rather than word regions. A future RapidOCR/Paddle provider should use a separately packaged persistent model host, pin and verify every model, expose word-level geometry where possible, and prove warm throughput before adoption.

PaddleOCR and RapidOCR are Apache-2.0 projects; RapidOCR deploys compact PaddleOCR-derived models over selectable inference engines. Direct PaddleOCR 3.x adds its own Python pipeline and a separately selected inference runtime, so RapidOCR/ONNX was the lower-friction neural candidate for this dependency-focused comparison. ONNX Runtime is MIT. Tesseract and the official `tessdata_fast` repository are Apache-2.0. These are compatible candidates, but redistribution would still require package and model notices; OfficeIMO currently leaves the native runtime external and downloads only explicitly requested, checksum-pinned Tesseract language data.

Primary references: [Tesseract](https://github.com/tesseract-ocr/tesseract), [tessdata_fast](https://github.com/tesseract-ocr/tessdata_fast), [RapidOCR](https://github.com/RapidAI/RapidOCR), [PaddleOCR installation](https://github.com/PaddlePaddle/PaddleOCR/blob/main/docs/version3.x/installation.en.md), and [ONNX Runtime](https://github.com/microsoft/onnxruntime).

## Recorded run

The checked-in evidence is from 2026-09-01 on Windows with Ubuntu 24.04.3 under WSL, an AMD Ryzen 9 9950X3D2, PowerShell 7.6.4, zero warmups, and three measured iterations. The full local PowerForge output remains ignored; the curated result is [`evidence/2026-09-01-windows-wsl-cpu.json`](evidence/2026-09-01-windows-wsl-cpu.json).

| Case | Tesseract median / CER | RapidOCR median / CER | Observation |
| --- | ---: | ---: | --- |
| Clean English | 571 ms / 0.000 | 6.47 s / 0.000 | Both exact |
| Dense table | 571 ms / 0.000 | 6.16 s / 0.010 | RapidOCR added one digit |
| Low contrast and 1.8-degree skew | 488 ms / 0.000 | 6.03 s / 0.000 | Both exact |
| Mixed English and Polish | 598 ms / 0.115 | 6.04 s / 0.034 | RapidOCR retained more diacritics |

The footprint recorded for the isolated test payload was 25.32 MiB for the extracted Tesseract packages and language data, versus 425.02 MiB for RapidOCR, ONNX Runtime, Python dependencies, and models. This is not a universal benchmark: persistent neural inference or GPU execution changes the latency profile, and a broader real-scan corpus is required before selecting a neural provider for a particular product.

## Run

Prepare an isolated environment under `Ignore/Benchmarks/OcrEngineComparison` with the layout and pinned versions in [`environment.lock.json`](environment.lock.json):

```text
tesseract-root/   extracted Ubuntu Tesseract packages
rapid-packages/   rapidocr and onnxruntime Python target directory
rapid-models/     downloaded ONNX models
```

No benchmark dependency belongs in the solution or a runtime project. After preparing that ignored environment:

```powershell
.\OfficeIMO.Reader.Ocr.Benchmarks.Comparisons\Invoke-OcrEngineComparison.ps1 -Plan
.\OfficeIMO.Reader.Ocr.Benchmarks.Comparisons\Invoke-OcrEngineComparison.ps1 -IterationCount 1 -Case clean-english
.\OfficeIMO.Reader.Ocr.Benchmarks.Comparisons\Invoke-OcrEngineComparison.ps1 -IterationCount 3
```

The wrapper uses PSPublishModule/PowerForge for matrix expansion, rotated ordering, measurement, validation, metrics, and JSON/CSV/Markdown artifacts. Successful lanes must emit non-empty text and positive geometry, and must remain below a 50% normalized character-error validity ceiling.
