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

### Working with checkboxes

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var para = document.AddParagraph("Accept:");
    para.AddCheckBox(false, "Consent", "ConsentTag");
    document.Save();
}

using (WordDocument document = WordDocument.Load(filePath)) {
    var cb = document.GetCheckBoxByTag("ConsentTag");
    cb.IsChecked = true;
    Console.WriteLine(cb.Alias);
}
```

### Combo boxes

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var items = new[] { "One", "Two" };
    document.AddParagraph().AddComboBox(items, "Combo", "ComboTag");
    document.Save();
}
```

### Picture controls

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddParagraph().AddPictureControl(imagePath, 100, 100, "Pic", "PicTag");
    document.Save();
}
```

### Repeating sections

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddParagraph().AddRepeatingSection("Section", "RS", "RSTag");
    document.Save();
}
```
