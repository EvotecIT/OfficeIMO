# OfficeIMO.OneNote.Html

`OfficeIMO.OneNote.Html` converts the typed offline OneNote model to HTML without Microsoft Graph, a OneNote installation, or a commercial dependency. It uses `OfficeIMO.OneNote.Markdown` as the shared semantic projection and the first-party `OfficeIMO.Markdown` HTML renderer.

```csharp
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Html;

OneNoteSection section = OneNoteSectionReader.Read("Section.one");
string html = section.ToHtmlDocument();
section.SaveAsHtml("Section.html");
```

Use `OneNoteMarkdownOptions` to include conflict copies or version-history pages and to resolve extracted asset destinations. HTML rendering remains fully offline unless a caller explicitly configures external assets in `HtmlOptions`.
