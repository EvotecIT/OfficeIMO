# README guide

How the root README and the package READMEs are structured. The root README is
the repository's front page; package READMEs are also the NuGet package pages
(`PackageReadmeFile`), so they must read well on nuget.org as well as on GitHub.

## What a README is for

A README tells someone deciding whether to use a package what it does, how to
install it and how to start. It is not the full contract and not a history.

| Content | Where it goes |
| --- | --- |
| What it does, install, first sample, main capabilities, links | README |
| Exact coverage, known limits, fidelity per feature | The package's support matrix or `SUPPORT.md` |
| Behaviour detail (option-by-option rules, edge cases, diagnostics codes) | The format guide in `Docs/` or the officeimo.com docs page |
| Measured performance | `Docs/performance/` and the officeimo.com benchmarks page |
| Upgrade steps | `MIGRATION.md` |
| Release history | GitHub Releases |
| Runtime differences between .NET Framework and modern .NET | One platform page in `Docs/`, linked from the root README |

Rules that apply to every README:

- Describe what works today, through measured workloads and the support
  matrices. No competitor comparisons, no PR progress, no "now supports".
- No hardcoded version numbers or format/route counts in prose. The NuGet badge
  shows the current version; counts come from generated data or link to it.
- Code samples are short and compile. Blocks wrapped in `powerforge:code`
  markers in READMEs listed in `Build/docs-sync.json` are generated from
  `Build/DocumentationExamples`; edit the example source and run the sync,
  never the marked block. Samples outside those markers are maintained by hand.
- Package READMEs use absolute links (`https://github.com/EvotecIT/OfficeIMO/blob/master/...`
  or `https://officeimo.com/...`). Relative links break on nuget.org. The root
  README is not packed and can keep relative links.
- Users install packages from NuGet. Don't send them to `Build/` or test
  fixture paths, and don't describe internal checks.

## Root README

1. **Title and badges**: CI, NuGet, PowerShell Gallery, license, Discord.
2. **What OfficeIMO is**: one paragraph. Open-source engines that create, edit,
   read and convert Office, PDF and many other document formats from .NET and
   PowerShell, without Microsoft Office or COM.
3. **Ways to use it**: a short table of the surfaces, each with its install
   line and link: .NET packages, PSWriteOffice, the `officeimo` CLI, the MCP
   server and agent plugin, OfficeIMO Studio, and the browser tools on
   officeimo.com.
4. **Why OfficeIMO**: five or six bullets, each backed by something a reader can
   check: first-party engines (dependency table in the package catalog),
   Windows, macOS and Linux with `netstandard2.0` and `net472` targets, NativeAOT
   and WebAssembly use (AOT smoke projects, browser tools), conversion reports
   that name what a target could not carry, bounded reads of untrusted input,
   and measured performance with a link to the published workloads.
5. **Quick start**: the generated QuickStart example.
6. **Formats**: a format map grouped by family (documents, spreadsheets,
   presentations, PDF and fixed layout, diagrams and plans, databases, notes
   and books, email and archives, markup and data, web and cloud). One row per
   format: extensions, package, and create / edit / read / convert marks.
   The marks come from `Website/data/format_map.json` and the generated operation
   contract, so the table should be generated rather than kept by hand.
7. **Read for search and AI**: Reader, OCR and the MCP tools in two or three
   lines.
8. **Packages**: how packages are named (`OfficeIMO.{Format}`,
   `OfficeIMO.{Source}.{Target}`, `OfficeIMO.Reader.{Format}`) and a link to the
   package catalog instead of a second full list.
9. **Platforms**: targets, operating systems, AOT, and one link to the platform
   differences page.
10. **Documentation, support, sponsors, license**.

## Engine package README

For a format engine (`OfficeIMO.Word`, `OfficeIMO.Pdf`, `OfficeIMO.Visio`, ...).
Aim for 100 to 200 lines.

```markdown
# OfficeIMO.{Format}

<NuGet version and downloads badges>

One or two sentences: what it creates, edits, reads or converts, which file
types, and that it needs no Office installation.

## Install

dotnet add package OfficeIMO.{Format}

Related packages: converters and readers for this format, one line each.

## Quick start

One sample of 10 to 25 lines that does the most common job.

## What it does

Capability bullets grouped by job: create, edit, read, convert, inspect and
secure. Each bullet names the entry point (type or method) rather than
repeating its full signature.

## Converting to other formats

Table: target, package, what carries over, where the loss report applies.

## Limits

Two to four sentences on the main boundaries, then a link to the support
matrix.

## Learn more

officeimo.com product and docs pages, API reference, examples,
support matrix, matching PSWriteOffice commands and CLI commands.

## Targets and dependencies

Target frameworks, OfficeIMO dependencies, external dependencies, license.

<generated capability summary block>
```

The `officeimo-operation-catalog` block at the end is written by
`Build/CompatibilityCatalog`; leave it in place and keep it last.

## Converter and reader package README

For `OfficeIMO.{Source}.{Target}`, `OfficeIMO.Reader.{Format}` and other
adapters. Aim for 40 to 80 lines.

```markdown
# OfficeIMO.{Source}.{Target}

<NuGet badges>

One sentence: which direction(s) it converts and the entry points
(`doc.SaveAsPdf(...)`, `pdf.ToWord(...)`).

## Install
## Quick start
## What carries over
Bullets for the content that maps, and what is reported as loss.
## Learn more
Support matrix, the two engine READMEs, the conversion guide on officeimo.com.
## Targets and dependencies
<generated capability summary block>
```

`OfficeIMO.Excel.Csv` is close to this shape today.

## Tool and app READMEs

`OfficeIMO.Tool`, `PowerShell/PSWriteOffice`, `OfficeIMO.Studio` and
`Apple/OfficeIMOStudio` keep the same opening (what it is, install, first
commands), then list command or feature groups and link to the generated
command reference (`PowerShell/PSWriteOffice/Docs`, `officeimo help`,
`Website/data/studio_features.json`) instead of documenting every command.
