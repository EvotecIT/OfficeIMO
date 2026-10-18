---
title: "Inspect PDF Signatures and Compliance"
description: "Collect signature structure, compliance, diagnostics, and rewrite-preservation evidence before accepting or transforming a PDF."
layout: docs
---

A signature object is not the same as a trusted signature decision. PSWriteOffice can expose signature structure and prepare or apply external signatures, but the surrounding workflow still owns trust stores, certificate policy, signer identity, timestamps, and acceptance rules.

## Inspect before rewriting

Use `Get-OfficePdfSignature`, `Get-OfficePdfCompliance`, `Get-OfficePdfDiagnostic`, and `Get-OfficePdfPreflight` to collect evidence. `Test-OfficePdfRewrite` helps assess whether a transformed candidate preserved the required document features.

## External signing flow

`New-OfficePdfSignature` prepares a signature placeholder and report. An external signing system produces the signature payload. `Set-OfficePdfSignature` applies that payload to the prepared file. Keep the exact prepared artifact and signing report together; changing the PDF between those steps invalidates the byte ranges.

For PDF/A, PDF/UA, electronic invoice, or other compliance workflows, inspect the specific compliance report and diagnostics relevant to the required standard. A readable PDF is not automatically compliant, accessible, or signature-preserving.

Start with [read and preflight](/docs/pswriteoffice/pdf-read-preflight/) and keep every transformation in a new output path so the evidence chain remains clear.
