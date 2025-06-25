## Cloning lists with numbering settings

This snippet demonstrates how `WordList.Clone` duplicates a list while keeping restart numbering and level overrides intact.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var list = document.AddList(WordListStyle.Headings111);
    list.RestartNumberingAfterBreak = true;
    list.Numbering.Levels[0].SetStartNumberingValue(5);

    list.AddItem("First");
    list.AddItem("Second");

    var clone = list.Clone();
    clone.AddItem("Third from clone");

    document.Save();
}
```
