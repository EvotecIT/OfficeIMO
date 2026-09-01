# OCR engine comparison

This opt-in suite compares free local OCR engines without adding them to the normal OfficeIMO restore graph. It uses generated, validated English, dense-table, and low-contrast/skew fixtures. Each lane recognizes the same PNG through a one-shot WSL process boundary, which matches the isolation model available to OfficeIMO providers.

## Support decision

OfficeIMO supports Tesseract as the default local engine. The reusable `IOfficeOcrEngine`, `OfficeOcrEnginePdfProvider`, and `IPdfOcrProvider` seams remain the supported route for caller-owned neural, native, or cloud engines.

RapidOCR remains a useful candidate, but it is not part of the default package graph because the tested Python and ONNX deployment was about 425 MiB, took roughly 6 seconds for every one-shot request, and returned line regions rather than word regions. A future RapidOCR/Paddle provider should use a separately packaged persistent model host, pin and verify every model, expose word-level geometry where possible, and prove warm throughput before adoption.

The pinned RapidOCR setup selects one recognition language model at a time. Mixed English-and-Polish evidence is therefore excluded from the parity table: comparing Tesseract's `eng+pol` request with RapidOCR's Polish-only model would not measure equivalent work. The runner rejects such a lane unless a genuinely equivalent multilingual RapidOCR model is pinned later.

PaddleOCR and RapidOCR are Apache-2.0 projects; RapidOCR deploys compact PaddleOCR-derived models over selectable inference engines. Direct PaddleOCR 3.x adds its own Python pipeline and a separately selected inference runtime, so RapidOCR/ONNX was the lower-friction neural candidate for this dependency-focused comparison. ONNX Runtime is MIT. Tesseract and the official `tessdata_fast` repository are Apache-2.0. These are compatible candidates, but redistribution would still require package and model notices; OfficeIMO currently leaves the native runtime external and downloads only explicitly requested, checksum-pinned Tesseract language data.

Primary references: [Tesseract](https://github.com/tesseract-ocr/tesseract), [tessdata_fast](https://github.com/tesseract-ocr/tessdata_fast), [RapidOCR](https://github.com/RapidAI/RapidOCR), [PaddleOCR installation](https://github.com/PaddlePaddle/PaddleOCR/blob/main/docs/version3.x/installation.en.md), and [ONNX Runtime](https://github.com/microsoft/onnxruntime).

## Recorded run

The checked-in evidence is from 2026-09-01 on Windows with Ubuntu 24.04.3 under WSL, an AMD Ryzen 9 9950X3D2, PowerShell 7.6.4, zero warmups, and three measured iterations. The full local PowerForge output remains ignored; the curated result is [`evidence/2026-09-01-windows-wsl-cpu.json`](evidence/2026-09-01-windows-wsl-cpu.json).

| Case | Tesseract median / CER | RapidOCR median / CER | Observation |
| --- | ---: | ---: | --- |
| Clean English | 571 ms / 0.000 | 6.47 s / 0.000 | Both exact |
| Dense table | 571 ms / 0.000 | 6.16 s / 0.010 | RapidOCR added one digit |
| Low contrast and 1.8-degree skew | 488 ms / 0.000 | 6.03 s / 0.000 | Both exact |

The footprint recorded for the isolated test payload was 25.32 MiB for the extracted Tesseract packages and language data, versus 425.02 MiB for RapidOCR, ONNX Runtime, Python dependencies, and models. Region counts and confidence scores are intentionally omitted from cross-engine results because the pinned Tesseract lane emits word geometry while RapidOCR emits line geometry, and their confidence contracts are engine-specific. This is not a universal benchmark: persistent neural inference or GPU execution changes the latency profile, and a broader real-scan corpus is required before selecting a neural provider for a particular product.

## Run

Prepare an isolated environment under `Ignore/Benchmarks/OcrEngineComparison` with the layout and pinned versions in [`environment.lock.json`](environment.lock.json):

```text
tesseract-root/   extracted Ubuntu Tesseract packages
rapid-packages/   rapidocr and onnxruntime Python target directory
rapid-models/     downloaded ONNX models
```

The exact PNG inputs and expected text are checked in under [`fixtures`](fixtures). `-RefreshFixtures` is a maintainer-only regeneration path and fails unless Pillow 11.0.0 and the pinned DejaVu font bytes match their recorded SHA-256 values.
Install the isolated Python payload from [`requirements.lock.txt`](requirements.lock.txt); the benchmark wrapper additionally requires its complete tree to match the digest in `environment.lock.json`, so version-equivalent but modified files fail closed.

```powershell
wsl.exe -- python3 -m pip install --no-compile --target /mnt/c/path/to/rapid-packages --requirement /mnt/c/path/to/requirements.lock.txt
```

No benchmark dependency belongs in the solution or a runtime project. After preparing that ignored environment:

```powershell
.\OfficeIMO.Reader.Ocr.Benchmarks.Comparisons\Invoke-OcrEngineComparison.ps1 -Plan
.\OfficeIMO.Reader.Ocr.Benchmarks.Comparisons\Invoke-OcrEngineComparison.ps1 -IterationCount 1 -Case clean-english
.\OfficeIMO.Reader.Ocr.Benchmarks.Comparisons\Invoke-OcrEngineComparison.ps1 -IterationCount 3
```

Before planning or measuring, the wrapper fails closed unless the live Tesseract, RapidOCR, and ONNX Runtime versions match `environment.lock.json`. The complete extracted Tesseract payload and RapidOCR Python payload—including source, bytecode, native dependencies, and symbolic-link targets—are covered by deterministic whole-tree SHA-256 digests. The locked install uses `--no-compile`, every Python invocation disables bytecode writes, and an unexpected cache file changes the digest. The checked-in fixture tree is hashed as a second immutable input boundary. Every ONNX model is verified separately by SHA-256 and models are never populated by the runner. The wrapper then uses PSPublishModule/PowerForge for matrix expansion, rotated ordering, measurement, validation, metrics, and JSON/CSV/Markdown artifacts. Successful lanes must emit non-empty text and positive geometry, and must remain below a 50% normalized character-error validity ceiling.
