---
title: "Read and Parse Markdown Content"
description: "Read Markdown as a document model and extract front matter, headings, nodes, and tables without fragile regular expressions."
layout: docs
---

Markdown is simple text, but structure-aware parsing is safer than ad hoc regular expressions when documents contain front matter, nested blocks, callouts, tables, task lists, or several Markdown dialects.

## Inspect semantic structures

```powershell
$headings = Get-OfficeMarkdownHeading -Path '.\Article.md'
$frontMatter = Get-OfficeMarkdownFrontMatter -Path '.\Article.md'
$tables = Get-OfficeMarkdownTable -Path '.\Article.md' -AsObject
$nodes = Get-OfficeMarkdownNode -Path '.\Article.md' -MaxDepth 3
```

Reader profiles and URL restrictions make parsing behavior explicit. Use heading level, text, anchor, node type, and depth filters to return only the structures the workflow needs.

The [inspect-content recipe](https://github.com/EvotecIT/OfficeIMO/blob/c1f0e5322ee8119ac229d463b8e42e1423a3c7f6/PowerShell/PSWriteOffice/Examples/Markdown/Recipe-Markdown-InspectContent.ps1) creates a document with front matter and a table, then reads those structures with dedicated commands.
