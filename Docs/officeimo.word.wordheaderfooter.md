# WordHeaderFooter

Namespace: OfficeIMO.Word

```csharp
public class WordHeaderFooter
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordHeaderFooter](./officeimo.word.wordheaderfooter.md)

## Properties

### **Paragraphs**

```csharp
public List<WordParagraph> Paragraphs { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **Tables**

```csharp
public List<WordTable> Tables { get; }
```

#### Property Value

[List&lt;WordTable&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **ParagraphsPageBreaks**

```csharp
public List<WordParagraph> ParagraphsPageBreaks { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **ParagraphsHyperLinks**

```csharp
public List<WordParagraph> ParagraphsHyperLinks { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **ParagraphsFields**

```csharp
public List<WordParagraph> ParagraphsFields { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **ParagraphsBookmarks**

```csharp
public List<WordParagraph> ParagraphsBookmarks { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **ParagraphsEquations**

```csharp
public List<WordParagraph> ParagraphsEquations { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **ParagraphsStructuredDocumentTags**

Provides a list of paragraphs that contain Structured Document Tags

```csharp
public List<WordParagraph> ParagraphsStructuredDocumentTags { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **ParagraphsImages**

Provides a list of paragraphs that contain Image

```csharp
public List<WordParagraph> ParagraphsImages { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **PageBreaks**

```csharp
public List<WordBreak> PageBreaks { get; }
```

#### Property Value

[List&lt;WordBreak&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **Images**

Exposes Images in their Image form for easy access (saving, modifying)

```csharp
public List<WordImage> Images { get; }
```

#### Property Value

[List&lt;WordImage&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **Bookmarks**

```csharp
public List<WordBookmark> Bookmarks { get; }
```

#### Property Value

[List&lt;WordBookmark&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **Fields**

```csharp
public List<WordField> Fields { get; }
```

#### Property Value

[List&lt;WordField&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **HyperLinks**

```csharp
public List<WordHyperLink> HyperLinks { get; }
```

#### Property Value

[List&lt;WordHyperLink&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **Equations**

```csharp
public List<WordEquation> Equations { get; }
```

#### Property Value

[List&lt;WordEquation&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **StructuredDocumentTags**

```csharp
public List<WordStructuredDocumentTag> StructuredDocumentTags { get; }
```

#### Property Value

[List&lt;WordStructuredDocumentTag&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

## Constructors

### **WordHeaderFooter()**

```csharp
public WordHeaderFooter()
```

## Methods

### **AddParagraph(String)**

```csharp
public WordParagraph AddParagraph(string text)
```

#### Parameters

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddParagraph()**

```csharp
public WordParagraph AddParagraph()
```

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddHyperLink(String, Uri, Boolean, String, Boolean)**

```csharp
public WordParagraph AddHyperLink(string text, Uri uri, bool addStyle, string tooltip, bool history)
```

#### Parameters

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`uri` Uri<br>

`addStyle` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

`tooltip` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`history` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddHyperLink(String, String, Boolean, String, Boolean)**

```csharp
public WordParagraph AddHyperLink(string text, string anchor, bool addStyle, string tooltip, bool history)
```

#### Parameters

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`anchor` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`addStyle` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

`tooltip` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`history` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddHorizontalLine(BorderValues, Nullable&lt;Color&gt;, UInt32, UInt32)**

```csharp
public WordParagraph AddHorizontalLine(BorderValues lineType, Nullable<Color> color, uint size, uint space)
```

#### Parameters

`lineType` BorderValues<br>

`color` [Nullable&lt;Color&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

`size` [UInt32](https://docs.microsoft.com/en-us/dotnet/api/system.uint32)<br>

`space` [UInt32](https://docs.microsoft.com/en-us/dotnet/api/system.uint32)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddBookmark(String)**

```csharp
public WordParagraph AddBookmark(string bookmarkName)
```

#### Parameters

`bookmarkName` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddField(WordFieldType, Nullable&lt;WordFieldFormat&gt;, Boolean)**

```csharp
public WordParagraph AddField(WordFieldType wordFieldType, Nullable<WordFieldFormat> wordFieldFormat, bool advanced)
```

#### Parameters

`wordFieldType` [WordFieldType](./officeimo.word.wordfieldtype.md)<br>

`wordFieldFormat` [Nullable&lt;WordFieldFormat&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

`advanced` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddTable(Int32, Int32, WordTableStyle)**

```csharp
public WordTable AddTable(int rows, int columns, WordTableStyle tableStyle)
```

#### Parameters

`rows` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

`columns` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

`tableStyle` [WordTableStyle](./officeimo.word.wordtablestyle.md)<br>

#### Returns

[WordTable](./officeimo.word.wordtable.md)<br>
