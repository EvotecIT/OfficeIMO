## Using Content Controls (Structured Document Tags)

`OfficeIMO` allows inserting and editing simple content controls.
The example below creates a document with a single control, assigns a tag and updates its text.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var sdt = document.AddStructuredDocumentTag("Sample text", "ExampleAlias", "ExampleTag");

    // change the text later
    sdt.Text = "Updated text";
    document.Save();
}

using (WordDocument document = WordDocument.Load(filePath)) {
    var loaded = document.GetStructuredDocumentTagByTag("ExampleTag");
    Console.WriteLine($"Tag text: {loaded.Text}");
}
```

### Multiple controls

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddStructuredDocumentTag("First", "Alias1", "Tag1");
    document.AddStructuredDocumentTag("Second", "Alias2", "Tag2");
    document.AddStructuredDocumentTag("Third", "Alias3", "Tag3");
    document.Save();
}

using (WordDocument document = WordDocument.Load(filePath)) {
    foreach (var control in document.StructuredDocumentTags) {
        Console.WriteLine(control.Tag + ": " + control.Text);
    }
}
```

### Advanced retrieval

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddStructuredDocumentTag("First", "Alias1", "Tag1");
    document.AddStructuredDocumentTag("Second", "Alias2", "Tag2");
    document.Save();
}

using (WordDocument document = WordDocument.Load(filePath)) {
    var alias = document.GetStructuredDocumentTagByAlias("Alias2");
    alias.Text = "Updated";
    var tag = document.GetStructuredDocumentTagByTag("Tag1");
    Console.WriteLine(tag.Text);
}
```
