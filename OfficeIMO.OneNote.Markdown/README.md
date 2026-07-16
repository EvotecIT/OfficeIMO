# OfficeIMO.OneNote.Markdown

`OfficeIMO.OneNote.Markdown` is the shared semantic projection from offline `OfficeIMO.OneNote` models to Markdown text and `MarkdownDoc`. Reader, HTML, and PDF adapters consume this projection instead of implementing separate OneNote format logic.

```csharp
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Markdown;

OneNoteSection section = OneNoteSectionReader.Read("Notes.one");
string markdown = section.ToMarkdown();
```

The package is MIT licensed and has no Graph, GraphEssentialsX, COM, or installed-OneNote dependency.
