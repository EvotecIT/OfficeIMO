## Using Content Controls (Structured Document Tags)

`OfficeIMO` allows inserting and editing simple content controls.
The example below creates a document with a single control and updates its text.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var sdt = document.AddStructuredDocumentTag("Sample text", "ExampleAlias");

    // change the text later
    sdt.Text = "Updated text";
    document.Save();
}
```
