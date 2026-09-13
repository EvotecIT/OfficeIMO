---
title: "Check and remove file provenance"
description: "Inspect Content Credentials and file-origin records in images and documents, then remove selected carrier categories from a separate browser-local copy."
meta.seo_title: "Check file provenance and Content Credentials | OfficeIMO"
layout: page
---

File provenance is origin or editing-history data stored inside a file. It can include **Content Credentials**, links to external credential manifests, and AI source declarations. OfficeIMO lets you inspect the supported records before deciding whether to keep them.

<div class="imo-intent-hero__actions">
  <a class="imo-btn imo-btn-primary" href="/convert/?workspace=provenance">Check a file in the browser</a>
  <a class="imo-btn imo-btn-ghost" href="https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.provenance-support-matrix.md">Read the provenance support matrix</a>
</div>

The browser tool accepts JPEG, PNG, WebP, PDF, DOCX, XLSX, and PPTX files up to 25 MB. The file stays in the current browser tab. Inspection does not upload it to an OfficeIMO server.

## What the browser tool does

1. Choose a supported image or document.
2. Inspect the file for supported origin records.
3. Review each finding and its location.
4. Choose which supported carrier categories to remove. Each selection removes every eligible record in that category.
5. Create a separate copy and inspect that copy again.
6. Download the copy and the optional JSON report.

Your original file is never overwritten. When a record is malformed, ambiguous, or unsafe to rewrite, OfficeIMO preserves or rejects it instead of silently damaging the file.

## What it can remove

- Embedded Content Credentials manifests.
- References to external credential manifests.
- AI-specific IPTC source declarations.

Removing these records can break an existing signature or credential chain. The result therefore explains what changed and re-inspects the generated copy before offering it for download.

## What it does not do

This tool does not remove visible watermarks, logos, text printed on a page, image pixels, or unrelated personal metadata. It does not fetch external credential data, validate signer trust, prove that a file is authentic, or decide whether a human or AI created it. A file with no detected Content Credentials is simply a file with no supported record found; that absence is not proof of origin.

Use the [provenance support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.provenance-support-matrix.md) when you need exact format and carrier coverage.
