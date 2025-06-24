# Custom Lists

OfficeIMO supports creating custom bullet lists using `AddCustomBulletList` and `AddCustomList`.
`AddCustomBulletList` creates a list with one bullet symbol for all levels, while `AddCustomList`
allows you to manually configure each level.

```csharp
var custom = document.AddCustomBulletList(WordBulletSymbol.Square, "Courier New", SixLabors.ImageSharp.Color.Red, fontSize: 16);
custom.AddItem("Custom bullet item");
```

```csharp
var builder = document.AddCustomList()
    .AddListLevel(1, WordBulletSymbol.Square, "Courier New", colorHex: "#FF0000", fontSize: 14)
    .AddListLevel(5, WordBulletSymbol.BlackCircle, "Arial", colorHex: "#00FF00", fontSize: 10);
builder.AddItem("First");
builder.AddItem("Fifth", 4);
```

See [`WordBulletSymbol`](./officeimo.word.wordbulletsymbol.md) for the available bullet symbols.

