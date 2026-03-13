# OfficeIMO.Markdown.Html

HTML to Markdown conversion for `OfficeIMO.Markdown`.

`OfficeIMO.Markdown.Html` is the HTML ingestion layer for the OfficeIMO markdown stack. It converts HTML fragments or full documents into:

- Markdown text
- `MarkdownDoc` block models from `OfficeIMO.Markdown`

The goal is not just "good looking output", but a structural conversion that keeps as much meaningful ordering and block shape as the current markdown AST allows.

## Design goals

- Convert HTML into a real `MarkdownDoc` first, then render Markdown text from that model.
- Preserve block ordering whenever HTML mixes paragraphs with quotes, nested lists, details blocks, and other supported structures.
- Resolve links and images consistently when a base URI is supplied.
- Preserve unsupported HTML explicitly when requested instead of silently flattening everything away.

## Current conversion model

Supported block-level mappings include:

- headings
- paragraphs
- ordered and unordered lists
- block quotes
- fenced code blocks from `pre` / `code`
- horizontal rules
- tables
- images and figures
- details / summary
- definition lists
- raw HTML fallback blocks for unsupported elements

Supported inline mappings include:

- emphasis and strong emphasis
- strike-through
- code spans
- links
- images
- hard line breaks
- a conservative set of inline passthrough elements that collapse to their children

## Usage

### Convert to Markdown text

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

var markdown = "<h1>Hello</h1><p>Body</p>".ToMarkdown();
var document = "<h1>Hello</h1><p>Body</p>".LoadFromHtml();
```

### Convert with options

```csharp
using OfficeIMO.Markdown.Html;

var options = new HtmlToMarkdownOptions {
    BaseUri = new Uri("https://example.com/docs/"),
    UseBodyContentsOnly = true,
    PreserveUnsupportedBlocks = true,
    PreserveUnsupportedInlineHtml = true
};

string markdown = "<p><a href=\"guide/start\">Docs</a></p>".ToMarkdown(options);
```

### Use the converter directly

```csharp
using OfficeIMO.Markdown.Html;

var converter = new HtmlToMarkdownConverter();
var document = converter.ConvertToDocument("<article><h1>Hello</h1><p>Body</p></article>");
```

## Options

- `BaseUri`
  Resolves relative link and image targets against a document base.
- `UseBodyContentsOnly`
  Uses `<body>` content when present instead of converting the whole HTML document node tree.
- `RemoveScriptsAndStyles`
  Drops `script`, `style`, `noscript`, and `template`.
- `PreserveUnsupportedBlocks`
  Emits unsupported block elements as `HtmlRawBlock` instead of dropping them.
- `PreserveUnsupportedInlineHtml`
  Emits unsupported inline elements as raw HTML instead of flattening them to plain text only.

## Structural notes

- Mixed block order inside list items is preserved.
- Multiple `dd` values for the same `dt` are preserved.
- Multiple `dt` terms sharing the same `dd` group are preserved.
- Unsupported custom/container elements are treated as block-level content when they are structurally block-like or when raw block preservation is enabled.
- Conversion happens through the `OfficeIMO.Markdown` AST, so the effective fidelity is bounded by that model.

## Current limitations

- Table cells are currently converted to Markdown cell text rather than a richer nested block model.
- Definition lists currently target the existing `OfficeIMO.Markdown` definition list representation.
- Unsupported HTML is preserved best when `PreserveUnsupportedBlocks` / `PreserveUnsupportedInlineHtml` are enabled.

## Related packages

- `OfficeIMO.Markdown`
  Core markdown AST, reader, and writer.
- `OfficeIMO.Reader.Html`
  HTML ingestion and chunking built on top of this converter.
