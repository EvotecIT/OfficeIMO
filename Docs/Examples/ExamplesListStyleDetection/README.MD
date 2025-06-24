## Detect list style for paragraphs

This example shows how to determine the style of a paragraph that belongs to a list.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var bullet = document.AddList(WordListStyle.Bulleted);
    bullet.AddItem("Bullet item");

    var numbered = document.AddList(WordListStyle.Headings111);
    numbered.AddItem("Numbered item");

    document.Save();
}

using (WordDocument document = WordDocument.Load(filePath)) {
    foreach (var paragraph in document.Paragraphs) {
        if (paragraph.IsListItem) {
            Console.WriteLine($"{paragraph.Text}: {paragraph.GetListStyle()}");
        }
    }
}
```
