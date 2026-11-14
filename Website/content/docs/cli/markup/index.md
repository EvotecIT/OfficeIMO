---
title: "OfficeIMO Markup from the CLI"
description: "Validate OfficeIMO Markup, generate C# or PowerShell from it, and export Word, Excel and PowerPoint files with officeimo markup."
layout: docs
---

OfficeIMO Markup is Markdown with explicit Office constructs: `@slide`, `@section` and `@sheet` containers and blocks such as `::notes`, `::chart`, `::range` and `::formula`. A file declares one profile, `document`, `workbook` or `presentation`, in its front matter. `officeimo markup` runs the same pipeline as the `OfficeIMO.Markup` packages; the [text formats guide](/docs/text-formats/) describes the language.

```markdown
---
profile: presentation
title: Quarterly Review
---

# Quarterly Review

@slide {
  layout: title-and-content
}

- Revenue grew 18%
- Churn improved

::notes
Open with the top-line result.
```

## Validate

```shell
officeimo markup validate deck.omd
```

```json
{
  "Diagnostics": [],
  "HasErrors": false
}
```

`validate` exits with code `1` when the markup has errors, so it can gate a pull request. Each diagnostic has a `Severity` and a `Message`, and names the block it refers to when there is one.

## Parse and preview

```shell
officeimo markup parse report.omd
officeimo markup parse notes.md --profile document
officeimo markup preview workbook.omd
```

`parse` and `preview` print the parsed document model and its diagnostics as one JSON object: the profile, the front-matter metadata and the blocks. Use `--profile presentation|document|workbook|common` for a file without a `profile` in its front matter.

## Generate code

```shell
officeimo markup emit deck.omd --target csharp --output BuildDeck.cs
officeimo markup emit workbook.omd --target powershell
```

`emit` writes C# that uses the OfficeIMO libraries, or PowerShell that uses [PSWriteOffice](/docs/pswriteoffice/). Without `--output` the code goes to standard output. Treat the result as a starting point you own and extend.

## Export Office files

```shell
officeimo markup export deck.omd --target pptx --output deck.pptx
officeimo markup export report.omd --target docx --output report.docx
officeimo markup export workbook.omd --target xlsx --output workbook.xlsx
```

```json
{
  "OutputPath": "/work/deck.pptx",
  "Target": "pptx"
}
```

`--target` must match the file's profile: `pptx` for presentations, `docx` for documents and `xlsx` for workbooks. A mismatch fails with exit code `5` and writes nothing. `markup emit --output` and `markup export` replace an existing file at the output path, so write to a new path when you need to keep the previous result.

Workbook export runs a safe preflight, repairs defined names and validates the Open XML package before saving. `--no-safe-preflight`, `--no-defined-name-repair` and `--no-openxml-validation` turn those steps off. Presentation export renders Mermaid diagrams through a renderer you configure with `--mermaid-renderer <mmdc>`; `--no-mermaid` skips them.

## Read from standard input

```shell
cat deck.omd | officeimo markup validate -
officeimo markup validate --stdin < deck.omd
```

`--max-input-bytes` bounds the markup size; the default is 64 MiB.

## Related

- [Text formats and OfficeIMO Markup](/docs/text-formats/)
- [Scripting and CI](/docs/cli/scripting-and-ci/)
