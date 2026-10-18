---
title: "Convert and Publish Markdown"
description: "Convert Markdown to styled HTML, convert HTML back to Markdown, and configure links and images for a safe publishing workflow."
layout: docs
---

Markdown is often the editable source while HTML is the delivery artifact. Keep that ownership clear: update Markdown, regenerate HTML, and avoid hand-editing both representations.

## Publish HTML

`ConvertTo-OfficeMarkdownHtml` supports fragment or document output, visual themes, CSS delivery modes, anchor links, task lists, footnotes, external-link attributes, and explicit URL or image restrictions.

```powershell
ConvertTo-OfficeMarkdownHtml -Path '.\Runbook.md' `
    -OutputPath '.\Runbook.html' -DocumentMode `
    -Title 'Operations runbook' -IncludeAnchorLinks
```

The [publish-HTML recipe](https://github.com/EvotecIT/PSWriteOffice/blob/main/Examples/Markdown/Recipe-Markdown-PublishHtml.ps1) creates a runbook and publishes a complete HTML document.

## Bring HTML back

Use `ConvertFrom-OfficeMarkdownHtml` for imported web fragments or articles. Base64 image handling, output directories, and URL policies should be chosen explicitly when the HTML is not fully trusted. Conversion preserves useful document meaning but is not expected to reproduce every browser layout detail.
