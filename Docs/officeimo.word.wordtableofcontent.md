# WordTableOfContent

Namespace: OfficeIMO.Word

```csharp
public class WordTableOfContent
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordTableOfContent](./officeimo.word.wordtableofcontent.md)

## Properties

### **Text**

```csharp
public string Text { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **TextNoContent**

```csharp
public string TextNoContent { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

## Constructors

### **WordTableOfContent(WordDocument, TableOfContentStyle)**

```csharp
public WordTableOfContent(WordDocument wordDocument, TableOfContentStyle tableOfContentStyle)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`tableOfContentStyle` [TableOfContentStyle](./officeimo.word.tableofcontentstyle.md)<br>

### **WordTableOfContent(WordDocument, SdtBlock)**

```csharp
public WordTableOfContent(WordDocument wordDocument, SdtBlock sdtBlock)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`sdtBlock` SdtBlock<br>

## Methods

### **Update()**

```csharp
public void Update()
```

### **Remove()**

```csharp
public void Remove()
```

### **Regenerate()**

```csharp
public WordTableOfContent Regenerate()
```

### Example

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddTableOfContent();
    document.AddParagraph("Heading 1").Style = WordParagraphStyles.Heading1;
    document.TableOfContent.Remove();
    document.RegenerateTableOfContent();
    document.Save();
}
```
