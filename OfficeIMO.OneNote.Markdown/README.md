# OfficeIMO.OneNote.Markdown

`OfficeIMO.OneNote.Markdown` is the shared semantic projection from offline `OfficeIMO.OneNote` models to Markdown text and `MarkdownDoc`. Reader, HTML, and PDF adapters consume this projection instead of implementing separate OneNote format logic.

```csharp
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Markdown;

OneNoteSection section = OneNoteSectionReader.Read("Notes.one");
string markdown = section.ToMarkdown();
```

Current pages are projected by default. Conflict copies and historical page versions are explicit because including them can duplicate or expose superseded content:

```csharp
var options = new OneNoteMarkdownOptions {
    IncludeConflictPages = true,
    IncludeVersionHistory = true
};

string markdownWithHistory = section.ToMarkdown(options);
```

Projection normalizes OneNote/RichEdit vertical-tab and form-feed paragraph separators to line breaks. Unsupported control characters, Unicode noncharacters, and unpaired surrogate code units are replaced with `?` so Markdown, HTML, and PDF consumers receive valid text. The original OneNote model is not mutated.

The package is MIT licensed and has no Graph, GraphEssentialsX, COM, or installed-OneNote dependency.
