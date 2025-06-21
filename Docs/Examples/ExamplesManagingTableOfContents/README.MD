## Managing Table of Contents

OfficeIMO.Word supports removing an existing table of contents and regenerating one in the same position. This can be useful when headings have changed and you want a fresh TOC.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddTableOfContent();
    document.AddParagraph("Heading 1").Style = WordParagraphStyles.Heading1;
    // remove the current table of contents
    document.TableOfContent.Remove();
    // recreate it at the original location
    document.RegenerateTableOfContent();
    document.Save();
}
```
