# WordBookmark

Namespace: OfficeIMO.Word

```csharp
public class WordBookmark
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordBookmark](./officeimo.word.wordbookmark.md)

## Properties

### **Name**

```csharp
public string Name { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Id**

```csharp
public int Id { get; set; }
```

#### Property Value

[Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

## Constructors

### **WordBookmark(WordDocument, Paragraph, BookmarkStart)**

```csharp
public WordBookmark(WordDocument document, Paragraph paragraph, BookmarkStart bookmarkStart)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`paragraph` Paragraph<br>

`bookmarkStart` BookmarkStart<br>

## Methods

### **Remove()**

```csharp
public void Remove()
```

### **AddBookmark(WordParagraph, String)**

```csharp
public static WordParagraph AddBookmark(WordParagraph paragraph, string bookmarkName)
```

#### Parameters

`paragraph` [WordParagraph](./officeimo.word.wordparagraph.md)<br>

`bookmarkName` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>
