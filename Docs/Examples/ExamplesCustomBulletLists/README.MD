## Custom bullet lists with colors

This example shows how to create bullet lists using `AddCustomBulletList` and set the bullet color and font.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var bulletList = document.AddCustomBulletList(WordListLevelKind.BulletDiamondSymbol, "Wingdings", Color.DarkRed, fontSize: 14);
    bulletList.AddItem("Red bullet item");
    bulletList.AddItem("Another item");

    var second = document.AddCustomBulletList('o', "Calibri", "0000ff", fontSize: 12);
    second.AddItem("Blue bullet item");
    second.AddItem("Second blue item");
    document.Save();
}
```
