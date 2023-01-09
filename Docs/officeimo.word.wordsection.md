# WordSection

Namespace: OfficeIMO.Word

```csharp
public class WordSection
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordSection](./officeimo.word.wordsection.md)

## Fields

### **Footer**

```csharp
public WordFooters Footer;
```

### **Header**

```csharp
public WordHeaders Header;
```

### **Borders**

```csharp
public WordBorders Borders;
```

### **Margins**

```csharp
public WordMargins Margins;
```

### **PageSettings**

```csharp
public WordPageSizes PageSettings;
```

## Properties

### **Paragraphs**

```csharp
public List<WordParagraph> Paragraphs { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **ParagraphsPageBreaks**

```csharp
public List<WordParagraph> ParagraphsPageBreaks { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **ParagraphsBreaks**

```csharp
public List<WordParagraph> ParagraphsBreaks { get; }
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

### **Breaks**

```csharp
public List<WordBreak> Breaks { get; }
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

### **Lists**

```csharp
public List<WordList> Lists { get; }
```

#### Property Value

[List&lt;WordList&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **Tables**

Provides a list of all tables within the section, excluding nested tables

```csharp
public List<WordTable> Tables { get; }
```

#### Property Value

[List&lt;WordTable&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **TablesIncludingNestedTables**

Provides a list of all tables within the section, including nested tables

```csharp
public List<WordTable> TablesIncludingNestedTables { get; }
```

#### Property Value

[List&lt;WordTable&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **DifferentFirstPage**

```csharp
public bool DifferentFirstPage { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **DifferentOddAndEvenPages**

```csharp
public bool DifferentOddAndEvenPages { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **PageOrientation**

```csharp
public PageOrientationValues PageOrientation { get; set; }
```

#### Property Value

PageOrientationValues<br>

### **ColumnsSpace**

```csharp
public Nullable<int> ColumnsSpace { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ColumnCount**

```csharp
public Nullable<int> ColumnCount { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

## Methods

### **GetType(String)**

```csharp
internal static HeaderFooterValues GetType(string type)
```

#### Parameters

`type` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

HeaderFooterValues<br>

### **ConvertRunsToWordParagraphs(WordDocument, Run)**

```csharp
internal static List<WordParagraph> ConvertRunsToWordParagraphs(WordDocument document, Run run)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`run` Run<br>

#### Returns

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **ConvertTableToWordTable(WordDocument, IEnumerable&lt;Table&gt;)**

```csharp
internal static List<WordTable> ConvertTableToWordTable(WordDocument document, IEnumerable<Table> tables)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`tables` [IEnumerable&lt;Table&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1)<br>

#### Returns

[List&lt;WordTable&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **ConvertParagraphsToWordParagraphs(WordDocument, IEnumerable&lt;Paragraph&gt;)**

```csharp
internal static List<WordParagraph> ConvertParagraphsToWordParagraphs(WordDocument document, IEnumerable<Paragraph> paragraphs)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`paragraphs` [IEnumerable&lt;Paragraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1)<br>

#### Returns

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **GetAllDocumentsLists(WordDocument)**

This method gets all lists for given document (for all sections)

```csharp
internal static List<WordList> GetAllDocumentsLists(WordDocument document)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

#### Returns

[List&lt;WordList&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **SetMargins(WordMargin)**

```csharp
public WordSection SetMargins(WordMargin pageMargins)
```

#### Parameters

`pageMargins` [WordMargin](./officeimo.word.wordmargin.md)<br>

#### Returns

[WordSection](./officeimo.word.wordsection.md)<br>

### **AddParagraph(String)**

```csharp
public WordParagraph AddParagraph(string text)
```

#### Parameters

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddWatermark(WordWatermarkStyle, String)**

```csharp
public WordWatermark AddWatermark(WordWatermarkStyle watermarkStyle, string text)
```

#### Parameters

`watermarkStyle` [WordWatermarkStyle](./officeimo.word.wordwatermarkstyle.md)<br>

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordWatermark](./officeimo.word.wordwatermark.md)<br>

### **SetBorders(WordBorder)**

```csharp
public WordSection SetBorders(WordBorder wordBorder)
```

#### Parameters

`wordBorder` [WordBorder](./officeimo.word.wordborder.md)<br>

#### Returns

[WordSection](./officeimo.word.wordsection.md)<br>

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
