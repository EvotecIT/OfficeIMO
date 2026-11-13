---
title: "Extract and Inspect Documents"
description: "Extract Markdown or JSON from documents and folders with officeimo read, inspect files as JSON and read workbook schemas."
layout: docs
---

The extraction commands turn documents into Markdown or structured JSON for search indexes, review tools, diffs and language models. They use OfficeIMO.Reader, which handles Office, OpenDocument, PDF, email and mailboxes, web formats, EPUB, Visio, OneNote, CSV, YAML, TOML and more.

## Extract one document

`officeimo read` and `officeimo extract` are the same command, a shortcut for `officeimo reader read`. Output goes to standard output unless you pass `--output`; an existing output file is refused unless you add `--force`:

```shell
officeimo read report.docx --format markdown
officeimo extract report.docx --format markdown --output report.md
officeimo extract handbook.pdf --format json --output handbook.json --assets handbook-assets
```

Read from standard input with `-`, and give the stream a name so the reader can pick the right handler:

```shell
cat report.docx | officeimo read - --name report.docx --format markdown
```

`--max-input-bytes` bounds the input size. `--assets <directory>` writes embedded images and other assets to a folder.

## Extract a folder

```shell
officeimo reader folder ./contracts --output ./contracts-md --format markdown
officeimo reader folder ./archive --output ./archive-json --format json --concurrency 4 --max-files 5000
```

Each source becomes `<file name>.md` or `<file name>.reader.json` in the output folder, for example `report.docx.md`. Existing files in the output folder are refused unless you add `--force`. Use `--no-recursive`, `--max-files`, `--max-total-bytes` and `--max-input-bytes` to bound the run. Apple `.pages`, `.numbers` and `.key` directory packages are read as single documents.

## See which formats are supported

```shell
officeimo reader capabilities
officeimo reader capabilities --format json
```

The text form prints one line per handler with its identifier and extensions. The command-line tool does not configure OCR or hosted providers, so scanned PDFs without a text layer return little text; see [searchable scans](/docs/cli/pdf/#searchable-scans-with-ocr).

## Reader JSON

JSON output, from `read --format json`, `reader folder --format json` or `convert <input> <output>.json` for a source without a data route, uses one versioned envelope. For YAML, TOML, XML, CSV and other data files, `convert` writes [converted data](/docs/cli/convert/#structured-data) instead, so use `read --format json` for their Reader JSON:

```json
{
  "schemaId": "officeimo.document.read-result",
  "schemaVersion": 15,
  "kind": "Word",
  "source": { "path": "...", "sourceId": "src:...", "sourceHash": "...", "lastWriteUtc": "...", "lengthBytes": 12478 },
  "markdown": "...",
  "chunks": [], "blocks": [], "tables": [], "pages": [], "metadata": [],
  "assets": [], "links": [], "forms": [], "visuals": [], "diagnostics": [], "nestedDocuments": []
}
```

Check `schemaId` and `schemaVersion` before parsing. `chunks` carry text, Markdown, a location, a token estimate and hashes for incremental indexing. `tables` carry columns, column profiles and rows, with `totalRowCount` and `truncated` when rows were limited. The [Reader guide](/docs/reader/) describes the model in detail.

## Inspect a file

`officeimo inspect` returns one compact JSON object, sized for a script or a model context rather than a full extraction:

```shell
officeimo inspect deck.pptx
officeimo inspect report.docx --max-output-characters 4000
```

```json
{
  "contentTrust": "untrusted",
  "contentSafety": { "status": "NotInspected" },
  "sourceId": "officeimo:f03752ed0b321dd92b95d277",
  "kind": "Word",
  "title": "", "author": "...",
  "preview": "Testing Tables on Load ...",
  "chunkCount": 4, "blockCount": 6, "pageCount": 0, "tableCount": 2,
  "metadata": [ { "name": "SectionCount", "value": "1" } ],
  "diagnostics": [],
  "truncated": false
}
```

`inspect` is a shortcut for `officeimo agent inspect`. The other `agent` commands search and fetch inside documents and mailboxes; see [MCP and agents](/docs/cli/mcp-and-agents/#agent-commands).

## Workbook sheets and schemas

```shell
officeimo tabular sheets workbook.xlsx
officeimo tabular schema workbook.xlsx --sheet Data
officeimo tabular schema export.csv --delimiter ';'
```

`sheets` prints one sheet name per line. `schema` prints a tab-separated `ordinal`, `name` and `type` table for each column, with types such as `string`, `integer`, `number`, `boolean` and `datetime`. Names escape backslashes and control characters as `\\`, `\t`, `\r`, `\n` or `\uXXXX`, so each stays on one line and is safe to split in a script.

## Mail data

```shell
officeimo agent inspect-email message.eml --max-output-characters 6000
officeimo agent search archive.pst --query "renewal date" --take 5
```

`inspect-email` reports body alternatives, charsets, attachment metadata, protection and signature status, and an HTML safety pass, without body previews. Mailbox search and fetch are covered in [MCP and agents](/docs/cli/mcp-and-agents/#agent-commands).

## Related

- [Convert documents](/docs/cli/convert/#markdown-and-json-output)
- [Reader guide](/docs/reader/)
- [Scripting and CI](/docs/cli/scripting-and-ci/)
