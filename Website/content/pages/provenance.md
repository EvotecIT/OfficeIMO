---
title: "Check and remove file provenance"
description: "Find Content Credentials and AI source labels in images and documents, then save a browser-local copy without the records you choose."
meta.seo_title: "Check file provenance and Content Credentials | OfficeIMO"
layout: conversion
meta.eyebrow: "Browser-local file inspection"
meta.source_format: "JPEG, PNG, WebP, PDF, DOCX, XLSX, or PPTX"
meta.destination_format: "Inspection report or cleaned copy"
meta.package: "OfficeIMO.Workflows + format packages"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Workflows"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_label: "Check a file in the browser"
meta.primary_url: "/browser/file-origin/"
meta.secondary_label: "Read the provenance support matrix"
meta.secondary_url: "https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.provenance-support-matrix.md"
meta.summary_title: "Inspection summary"
meta.limit: "One supported file up to 25 MB. Structural inspection does not prove origin, authenticity, or signer trust."
meta.related_label: "Browse all browser tools"
meta.related_url: "/convert/"
---

File provenance is origin or editing-history data stored inside a file. It can include **Content Credentials**, links to external credential manifests, and AI source declarations. OfficeIMO lets you inspect the supported records before deciding whether to keep them.

The browser tool accepts JPEG, PNG, WebP, PDF, DOCX, XLSX, and PPTX files up to 25 MB. The file stays in the current browser tab. Inspection does not upload it to an OfficeIMO server.

## What the browser tool does

1. Choose a supported image or document. The check starts straight away.
2. Review what was found: each kind of record, what it says, and where it is stored. For Content Credentials this includes what the record itself claims, described below.
3. Untick anything you want to keep. Each ticked kind removes every eligible record of that kind.
4. Save a separate copy. OfficeIMO inspects the copy again and lists what was removed, what was kept, and anything it could not remove.
5. Download the copy and, if you need it, the JSON report.

Your original file is never overwritten. When a record is malformed, ambiguous, or unsafe to rewrite, OfficeIMO preserves or rejects it instead of silently damaging the file.

## What Content Credentials say

When a file carries Content Credentials (C2PA), the tool reads the active record and shows what it claims:

- **Made with:** the app that wrote the record, such as an image service or editor, and the AI model named for each step when one is recorded.
- **AI-generated:** whether a step declares a generative AI source (the IPTC "trained algorithmic media" type) or a mix of AI and other content.
- **Steps:** what the record says happened, such as created, edited, converted, or cropped.
- **Based on:** other files the record lists as ingredients.
- **Signed by:** the organization named in the signing certificate, and who issued that certificate.

These are the record's own statements. The tool reads the certificate names but does not verify the signature, so it cannot tell whether the record was changed after it was written. Some records also say a watermark was added to the content itself; removing the record does not remove that watermark, and the result says so.

In .NET, the same summary is available as `OfficeProvenanceEvidence.Manifest` on each Content Credentials record returned by the provenance inspection API.

## Find hidden characters in text

The [hidden characters tool](/browser/hidden-characters/) reviews text the same way. Paste text or open a UTF-8 file, or a UTF-16/32 file with a byte-order mark. The text stays in the browser. This review accepts up to 1 MiB of encoded data, 1,048,576 UTF-16 code units and 512 findings.

Inspect the code-point markers, offsets and contextual risks, then select individual occurrences to remove from a separate text copy. No occurrences are selected automatically. Language joiners, emoji selectors and typographic spaces may be intentional; the tool does not normalize them or treat them as proof of AI authorship. File exports preserve the original encoding, byte-order mark and line endings. Editing the source clears the previous review and downloads.

The text report records exact findings and selected occurrence indices under `officeimo.text-integrity.result.v1`. Provenance reports use the same `officeimo.provenance.result.v2` contract as the CLI and Studio, with source/output hashes, coverage and explicit check states.

## What it can remove

- Embedded Content Credentials manifests.
- References to external credential manifests.
- AI-specific IPTC source declarations.

Removing these records can break an existing signature or credential chain. The result therefore explains what changed and re-inspects the generated copy before offering it for download.

## What it does not do

This tool does not remove visible or invisible watermarks, logos, text printed on a page, image pixels, or unrelated personal metadata. It does not fetch external credential data, validate signer trust, prove that a file is authentic, or decide whether a human or AI created it. A file with no detected Content Credentials is simply a file with no supported record found; that absence is not proof of origin.

Use the [provenance support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.provenance-support-matrix.md) when you need exact format and carrier coverage.
