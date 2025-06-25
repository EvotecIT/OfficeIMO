## Lists with picture bullets

Small example:

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    var pictureList = document.AddPictureBulletList(imagePath);
    pictureList.AddItem("First");
    pictureList.AddItem("Second");
    document.Save();
}
```

Using a stream:

```csharp
using var stream = File.OpenRead(imagePath);
using (WordDocument document = WordDocument.Create(filePath)) {
    var list = document.AddPictureBulletList(stream, "bullet.jpg");
    list.AddItem("Stream item");
    document.Save();
}
```

Bigger document:

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddParagraph("Shopping list:");
    var pictureList = document.AddPictureBulletList(imagePath);
    pictureList.AddItem("Milk");
    pictureList.AddItem("Bread");
    pictureList.AddItem("Butter");
    document.AddParagraph("End of list.");
    document.Save();
}
```
