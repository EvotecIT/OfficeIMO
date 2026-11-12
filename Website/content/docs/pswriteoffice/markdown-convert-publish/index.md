---
title: "Convert and Publish Markdown"
description: "Convert Markdown to styled HTML, convert HTML back to Markdown, and configure links and images for a safe publishing workflow."
layout: docs
---

Markdown is often the editable source while HTML is the delivery artifact. Keep that ownership clear: update Markdown, regenerate HTML, and avoid hand-editing both representations.

## Publish HTML

`Convert-OfficeDocument -To Html` with `New-OfficeMarkdownHtmlOptions` supports fragment or document output, visual themes, CSS delivery modes, anchor links, task lists, footnotes, external-link attributes, and explicit URL or image restrictions.

```powershell
$options = New-OfficeMarkdownHtmlOptions -DocumentMode -Title 'Operations runbook' -IncludeAnchorLinks
Convert-OfficeDocument -Path '.\Runbook.md' -To Html `
    -OutputPath '.\Runbook.html' -Options $options
```

The [publish-HTML recipe](https://github.com/EvotecIT/OfficeIMO/blob/15d71340482668a9cf0e669dfc6ee0ea0bfdbf99/PowerShell/PSWriteOffice/Examples/Markdown/Recipe-Markdown-PublishHtml.ps1) creates a runbook and publishes a complete HTML document.

## Bring HTML back

Use `Convert-OfficeDocument -From Html -To Markdown` for imported web fragments or articles. `New-OfficeHtmlMarkdownOptions` supplies the image policy, table limits and Markdown writer settings. Start with `-Portable` for portable Markdown, or use `-BaseOptions` to adjust a copy of existing native settings.
