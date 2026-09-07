---
title: "PDF corpus: inspect the files behind the results"
description: "Explore a measured OfficeIMO PDF corpus with source hashes, licenses, rendering results, extraction checks, mutation decisions, failures, and reproduction commands."
layout: page
meta.seo_title: "OfficeIMO PDF corpus and reproducible quality evidence"
---

These results come from the existing OfficeIMO PDF quality runner. Each input has a pinned source and SHA-256 hash. The runner opens it, exercises public APIs, and compares the outcome with explicit expectations.

A passing case means those expectations were met. It can include a safe refusal to modify a file. A rendered page, extracted text, or declared PDF/A marker does not establish visual fidelity, reading-order accuracy, or standards conformance.

{{< pdf-corpus >}}

## Other evidence answers different questions

| Evidence | What it establishes | What it does not establish |
|---|---|---|
| [PDF performance workloads](/benchmarks/) | Validated equivalent work and measurements for the recorded environment | Performance on every document or machine |
| [Office producer references](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Pdf.Tests/Pdf/ReferenceBaselines/reference-corpus.json) | Pinned Microsoft Office output, rendering distances, and declared budgets | Pixel equivalence or universal Office compatibility |
| [PDF interoperability input catalog](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Pdf.Benchmarks.Comparisons/Corpus/pdf-corpus.json) | Reproducible input selection with provenance and licensing | A completed run for every listed input |
| [Reverse conversion scorecard](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/pdf-reverse-conversion-scorecard.json) | Route-specific serialization, reopening, and semantic or visual contracts | Reconstruction of the original editable source |

The catalogs can overlap. Their entry counts must not be added together as a unique-file total. OCR accuracy, certified conformance, and large external-corpus stress results are separate evidence requirements.

Try a [PDF task](/pdf/), inspect a [generated document](/showcase/?format=pdf), or read the [runner and interpretation guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.real-world-corpus-evidence.md).
