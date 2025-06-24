# Custom Lists

OfficeIMO supports creating custom bullet lists using `AddCustomBulletList` and `AddCustomList`.
`AddCustomBulletList` creates a list with one bullet symbol for all levels, while `AddCustomList`
allows you to manually configure each level.

```csharp
var custom = document.AddCustomBulletList(WordListLevelKind.BulletSquareSymbol, "Courier New", SixLabors.ImageSharp.Color.Red, fontSize: 16);
custom.AddItem("Custom bullet item");
```

```csharp
var builder = document.AddCustomList()
    .AddListLevel(1, WordListLevelKind.BulletSquareSymbol, "Courier New", colorHex: "#FF0000", fontSize: 14)
    .AddListLevel(5, WordListLevelKind.BulletBlackCircle, "Arial", colorHex: "#00FF00", fontSize: 10);
builder.AddItem("First");
builder.AddItem("Fifth", 4);
```

You can adjust where numbering begins using `StartNumberingValue` on a level:

```csharp
var numbered = document.AddCustomList();
var level = new WordListLevel(WordListLevelKind.Decimal)
    .SetStartNumberingValue(3);
numbered.Numbering.AddLevel(level);
numbered.AddItem("Starts at three");
```

See [`WordListLevelKind`](./officeimo.word.wordlistlevelkind.md) for the available bullet symbols.

