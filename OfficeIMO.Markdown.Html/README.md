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
- shared `data-omd-*` visual host elements back into semantic fenced blocks
- raw HTML fallback blocks for unsupported elements

Supported inline mappings include:

- emphasis and strong emphasis
- strike-through
- code spans
- links
- images
- hard line breaks
- typed inline HTML wrappers for `q`, `u`, `ins`, `sub`, and `sup`
- a conservative raw/passthrough fallback for unsupported inline HTML when preservation is enabled

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

### Convert with portable markdown output

```csharp
using OfficeIMO.Markdown.Html;

var options = HtmlToMarkdownOptions.CreatePortableProfile();

string markdown = """
<blockquote>
  <p><strong>Example</strong></p>
  <p>Body text</p>
</blockquote>
""".ToMarkdown(options);
```

Use the portable profile when HTML ingestion should produce generic markdown output instead of OfficeIMO-specific block syntax.

### Convert to `MarkdownDoc`, then choose the markdown writer profile explicitly

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

var converter = new HtmlToMarkdownConverter();
var document = converter.ConvertToDocument("""
<table>
  <tr><th>Name</th><th>Notes</th></tr>
  <tr><td>Alice</td><td><p>Line one</p><blockquote><p>Line two</p></blockquote></td></tr>
</table>
""");

var officeMarkdown = document.ToMarkdown(MarkdownWriteOptions.CreateOfficeIMOProfile());
var portableMarkdown = document.ToMarkdown(MarkdownWriteOptions.CreatePortableProfile());
```

This is the cleanest path when HTML ingestion fidelity matters first and the markdown serialization contract is a separate downstream decision.

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
- `MarkdownWriteOptions`
  Controls how the intermediate `MarkdownDoc` is serialized back to markdown text.
  Use `HtmlToMarkdownOptions.CreatePortableProfile()` when portability matters more than preserving OfficeIMO-style output.

## Profile guidance

- `CreateOfficeIMOProfile()`
  Best when the downstream consumer is `OfficeIMO.Markdown`/`OfficeIMO.MarkdownRenderer` and can benefit from richer OfficeIMO block syntax.
- `CreatePortableProfile()`
  Best when the downstream consumer is a generic markdown engine, HTML reconversion flow, or another parser that should not depend on OfficeIMO-only syntax.

The important split is:

- `HtmlToMarkdownOptions`
  Controls HTML ingestion behavior and preservation choices.
- `MarkdownWriteOptions`
  Controls how the intermediate AST is written back to markdown text.

That means `OfficeIMO.Markdown.Html` is no longer just a text flattener. It is an HTML-to-AST bridge with a configurable markdown writer on the output side.

## Structural notes

- Mixed block order inside list items is preserved.
- Multiple `dd` values for the same `dt` are preserved.
- Multiple `dt` terms sharing the same `dd` group are preserved.
- Block-rich `dd` values are preserved as typed block content instead of being forced through inline-only conversion.
- Table cells preserve typed block content in the intermediate `MarkdownDoc` AST instead of collapsing immediately to strings.
- Supported inline HTML such as `q`, `u`, `ins`, `sub`, and `sup` is preserved as typed AST wrappers instead of being flattened to plain text.
- Unsupported custom/container elements are treated as block-level content when they are structurally block-like or when raw block preservation is enabled.
- Shared renderer visual hosts that carry the `data-omd-*` contract are decoded back into `SemanticFencedBlock` nodes, which lets `OfficeIMO.MarkdownRenderer` HTML round-trip into semantic markdown fences.
- Conversion happens through the `OfficeIMO.Markdown` AST, so the effective fidelity is bounded by that model.

For the current stack, this means HTML ingestion can preserve more structure than plain markdown text can always express directly. The AST is the source of truth; markdown emission is the profile-driven projection of that model.

## Current limitations

- Markdown text emission is still constrained by markdown syntax itself, so rich table-cell and definition-list AST content may be flattened when serialized for engines that only accept plain markdown text.
- Downstream converters may still choose deliberate degradations for AST-preserved HTML wrappers when the target format has no native equivalent. For example, the Word converter keeps `u/sub/sup` structurally but intentionally degrades `ins` and `q`.
- Portable output intentionally degrades OfficeIMO-specific constructs instead of preserving host-specific syntax.
- Unsupported HTML is preserved best when `PreserveUnsupportedBlocks` / `PreserveUnsupportedInlineHtml` are enabled.

## Related packages

- `OfficeIMO.Markdown`
  Core markdown AST, reader, and writer.
- `OfficeIMO.Reader.Html`
  HTML ingestion and chunking built on top of this converter.
