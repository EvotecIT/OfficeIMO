# OfficeIMO.Latex.Markdown

This package maps the bounded `OfficeIMO.Latex` profile to and from `OfficeIMO.Markdown`. It reports source fallbacks and simplifications, especially for TeX math layout, package-specific commands, bibliography formatting, and unknown environments.

```csharp
using OfficeIMO.Latex;
using OfficeIMO.Latex.Markdown;

LatexDocument latex = LatexDocument.Parse(source).Document;
LatexMarkdownConversionResult converted = latex.ToMarkdownDocument();
string markdown = converted.Document.ToMarkdown();
```

Reverse conversion creates canonical bounded-profile LaTeX and reparses the generated source through the lossless engine:

```csharp
MarkdownLatexConversionResult generated = markdownDocument.ToLatexDocument();
string source = generated.Source;
LatexDocument parsed = generated.Document;
```

The bridge maps front matter, headings, inline formatting and links, lists and definitions, images/figures, table captions/labels and common spans, theorem callouts with required declarations, verbatim/code, and math transport. Canonical output escapes TeX arguments and deterministically encodes labels. Unrepresented figure/table container source remains visible with diagnostics. It does not promise TeX layout or execute package behavior.
