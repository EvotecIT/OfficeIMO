## Updating fields programmatically

`WordDocument` allows two approaches to refresh fields like page numbers or table of contents.

- `Settings.UpdateFieldsOnOpen` instructs Word to update fields when the document is opened.
- `UpdateFields()` updates page numbers immediately and ensures the table of contents will refresh on open.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.Settings.UpdateFieldsOnOpen = true; // ask Word to refresh fields

    document.AddParagraph("Page 1").AddPageNumber(includeTotalPages: true);
    document.AddPageBreak();
    document.AddParagraph("Page 2");
    document.AddTableOfContent();

    document.UpdateFields();
    document.Save();
}
```
