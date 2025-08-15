# OfficeIMO Converters - Master TODO

## ✅ Current Implementation Status

### Core
- Project builds and compiles
- Extension methods work (`string.LoadFromHtml()`, `string.LoadFromMarkdown()`)
- Converters for HTML, Markdown and PDF share common infrastructure
- Basic null checking and exception handling

### HTML Converters
- HTML to Word handles paragraphs, headings (with Word styles), bold/italic/underline, hyperlinks, lists, tables (row/colspan), images (including SVG and base64), line breaks, blockquotes, code blocks, footnotes, CSS classes and inline styles, external stylesheets, and nested structures.
- Word to HTML exports headings, paragraph formatting, bold/italic/underline/strike, hyperlinks, lists, tables (with spans), images (external or embedded), footnotes, blockquotes, code blocks, metadata and optional base64 images.

### Markdown Converters
- Markdown to Word uses Markdig and supports headings, bold/italic, links, lists, code blocks, tables, images, blockquotes and footnotes.
- Word to Markdown exports headings, formatting, links, lists, tables, images and footnotes.

### PDF Converter
- Partially working: supports paragraphs, tables, lists and headers/footers.

### Known Limitations
- Advanced CSS layout (float, flex, grid) is not fully mapped.
- HTML forms, audio/video, canvas and script elements are ignored.
- Comments, revisions, shapes, charts and other complex Word features are not converted.
- PDF converter lacks section awareness and advanced element types.

## 📁 Project Structure

```
OfficeIMO/
├── OfficeIMO.Word/              # Core Word library
├── OfficeIMO.Word.Pdf/          # PDF converter
├── OfficeIMO.Word.Markdown/     # Markdown converter
├── OfficeIMO.Word.Html/         # HTML converter
├── OfficeIMO.Examples/          # Examples
└── OfficeIMO.Tests/             # Tests
```

## 🔧 Phase 1: HTML Converter Improvements

### HTML to Word Converter
#### Currently Implemented
- AngleSharp HTML parsing
- Paragraph and heading support with Word styles
- Bold, italic, underline, strike, superscript, subscript
- Hyperlinks and bookmarks
- Ordered and unordered lists with nesting
- Tables with captions, row/col spans and styles
- Images (embedded, external, SVG)
- Line breaks and horizontal rules
- Blockquotes, code blocks, abbreviations, footnotes and citations
- Inline and external CSS (with class mapping and basic property support)
- Page settings from options

#### Remaining Gaps
- HTML forms (`input`, `select`, `textarea`)
- Audio, video and canvas elements
- Advanced CSS layout (float, flexbox, grid, positioning)
- Style mapping for more CSS properties
- Performance optimization for large documents

### Word to HTML Converter
#### Currently Implemented
- Basic HTML document structure and UTF-8 meta
- Metadata export (title, author, keywords, etc.)
- Heading detection and `<h1>-<h6>` output
- Paragraph formatting (bold, italic, underline, strike, superscript, subscript)
- Hyperlinks and bookmarks
- Ordered and unordered lists with nesting
- Tables with row/col spans and captions
- Images with optional base64 embedding and SVG support
- Footnotes and citations
- Blockquotes, code blocks and horizontal rules
- Optional font-family override

#### Remaining Gaps
- Export comments, revisions and other annotations
- Shapes, charts, SmartArt and other drawing elements
- Advanced CSS style generation
- Form field export
- Better handling of custom styles and themes

## 🔧 Phase 2: Markdown Converter Enhancements

### Markdown to Word Converter
- **Implemented:** Uses Markdig; supports headings, bold, italic, links, lists, tables, images, blockquotes, code blocks, footnotes and horizontal rules.
- **Next Steps:** Task list items, table alignment options, math support.

### Word to Markdown Converter
- **Implemented:** Exports headings, bold, italic, links, lists, tables, images, footnotes and horizontal rules.
- **Next Steps:** Preserve custom styles, support comments/revisions, improved image handling.

## 📊 Real Progress Tracking

### Markdown Converter
- [x] Parse with Markdig
- [x] Headings with styles
- [x] Bold/Italic
- [x] Lists
- [x] Links
- [x] Code blocks
- [x] Tables
- [x] Images
- [ ] Task list items
- [ ] Math extension

### HTML Converter
- [x] Headings with styles
- [x] Bold/Italic/Underline/Strike
- [x] Hyperlinks
- [x] Lists
- [x] Tables
- [x] Images
- [x] CSS styles (basic)
- [ ] Forms and media elements
- [ ] Advanced CSS layout

### Word to HTML
- [x] Detect heading styles
- [x] Export formatting
- [x] Export hyperlinks
- [x] Export lists
- [x] Export tables
- [x] Export images
- [ ] Export comments and revisions
- [ ] Export shapes/charts

### Word to Markdown
- [x] Detect heading styles
- [x] Export formatting
- [x] Export hyperlinks
- [x] Export lists
- [x] Export tables
- [x] Export images
- [ ] Export comments and revisions

## 🚀 Next Steps
1. Extend HTML converters to handle forms, multimedia and advanced CSS layout.
2. Improve round-trip fidelity for comments, revisions and shapes.
3. Expand test coverage and add more real-world examples.
4. Optimize performance for large documents and heavy CSS usage.

## 🏁 Success Criteria

```
string markdown = "# Heading\n\n**Bold** and *italic* and [link](http://example.com)";
var doc = markdown.LoadFromMarkdown();
string html = doc.ToHtml();
string markdownOut = doc.ToMarkdown();

// Round trip assertions
Assert.That(markdownOut.Contains("# Heading"));
Assert.That(markdownOut.Contains("**Bold**"));
Assert.That(markdownOut.Contains("*italic*"));
Assert.That(markdownOut.Contains("[link](http://example.com)"));
```

When these round trips preserve formatting and structure across HTML, Markdown and Word documents, the converters can be considered feature-complete for basic scenarios.

