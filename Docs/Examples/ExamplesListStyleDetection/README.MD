## Detect list style for paragraphs

This example shows how to determine the style of a paragraph that belongs to a list
and how to read the style directly from a list instance.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var bullet = document.AddList(WordListStyle.Bulleted);
    bullet.AddItem("Bullet item");
    Console.WriteLine($"Bullet list uses style: {bullet.Style}");

    var numbered = document.AddList(WordListStyle.Headings111);
    numbered.AddItem("Numbered item");
    Console.WriteLine($"Numbered list uses style: {numbered.Style}");

    document.Save();
}

using (WordDocument document = WordDocument.Load(filePath)) {
    foreach (var list in document.Lists) {
        Console.WriteLine($"Loaded list style: {list.Style}");
    }
    foreach (var paragraph in document.Paragraphs) {
        if (paragraph.IsListItem) {
            Console.WriteLine($"{paragraph.Text}: {paragraph.ListStyle}");
        }
    }
}
```
