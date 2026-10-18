---
title: "Convert Between Markdown and Word"
description: "Create reviewable Word documents from Markdown and return DOCX content to Markdown with explicit image and formatting policies."
layout: docs
---

A Markdown-to-Word workflow is useful when authors prefer text-based source control but reviewers need comments, tracked changes, or a familiar DOCX file. The return conversion is best treated as semantic round-trip, not pixel-perfect fidelity.

## Create the review copy

`ConvertFrom-OfficeWordMarkdown` accepts Markdown text, a file path, or a parsed Markdown document. It can use a template, target a bookmark or content control, render front matter, and control local, remote, and data-URI images.

## Return to Markdown

`ConvertTo-OfficeWordMarkdown` converts a Word path or loaded document to Markdown. Image export mode and image directory determine whether embedded images become sidecar files. Underline and highlight output are opt-in because not every Markdown target represents them consistently.

The [Word round-trip recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Markdown/Recipe-Markdown-WordRoundTrip.ps1) creates Markdown, converts it to DOCX, and converts the Word file back to Markdown.

For heavily designed Word templates, define which semantic elements must survive before promising a round-trip. Tables, headings, lists, links, and basic emphasis are a better contract than exact pagination.
