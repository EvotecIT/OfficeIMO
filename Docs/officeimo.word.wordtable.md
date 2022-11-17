# WordTable

Namespace: OfficeIMO.Word

```csharp
public class WordTable
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordTable](./officeimo.word.wordtable.md)

## Fields

### **Position**

```csharp
public WordTablePosition Position;
```

## Properties

### **Paragraphs**

```csharp
public List<WordParagraph> Paragraphs { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **Style**

```csharp
public Nullable<WordTableStyle> Style { get; set; }
```

#### Property Value

[Nullable&lt;WordTableStyle&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Alignment**

```csharp
public Nullable<TableRowAlignmentValues> Alignment { get; set; }
```

#### Property Value

[Nullable&lt;TableRowAlignmentValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **WidthType**

```csharp
public Nullable<TableWidthUnitValues> WidthType { get; set; }
```

#### Property Value

[Nullable&lt;TableWidthUnitValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Width**

Gets or sets width of a table

```csharp
public Nullable<int> Width { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **LayoutType**

Gets or sets layout of a table

```csharp
public Nullable<TableLayoutValues> LayoutType { get; set; }
```

#### Property Value

[Nullable&lt;TableLayoutValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ConditionalFormattingFirstRow**

Specifies that the first row conditional formatting shall be applied to the table.

```csharp
public Nullable<bool> ConditionalFormattingFirstRow { get; set; }
```

#### Property Value

[Nullable&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ConditionalFormattingLastRow**

Specifies that the last row conditional formatting shall be applied to the table.

```csharp
public Nullable<bool> ConditionalFormattingLastRow { get; set; }
```

#### Property Value

[Nullable&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ConditionalFormattingFirstColumn**

Specifies that the first column conditional formatting shall be applied to the table.

```csharp
public Nullable<bool> ConditionalFormattingFirstColumn { get; set; }
```

#### Property Value

[Nullable&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ConditionalFormattingLastColumn**

Specifies that the last column conditional formatting shall be applied to the table.

```csharp
public Nullable<bool> ConditionalFormattingLastColumn { get; set; }
```

#### Property Value

[Nullable&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ConditionalFormattingNoHorizontalBand**

Specifies that the horizontal banding conditional formatting shall not be applied to the table.

```csharp
public Nullable<bool> ConditionalFormattingNoHorizontalBand { get; set; }
```

#### Property Value

[Nullable&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ConditionalFormattingNoVerticalBand**

Specifies that the vertical banding conditional formatting shall not be applied to the table.

```csharp
public Nullable<bool> ConditionalFormattingNoVerticalBand { get; set; }
```

#### Property Value

[Nullable&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **RowsCount**

```csharp
public int RowsCount { get; }
```

#### Property Value

[Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### **Rows**

```csharp
public List<WordTableRow> Rows { get; }
```

#### Property Value

[List&lt;WordTableRow&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **FirstRow**

```csharp
public WordTableRow FirstRow { get; }
```

#### Property Value

[WordTableRow](./officeimo.word.wordtablerow.md)<br>

### **LastRow**

```csharp
public WordTableRow LastRow { get; }
```

#### Property Value

[WordTableRow](./officeimo.word.wordtablerow.md)<br>

### **Title**

Gets or sets a Title/Caption to a Table

```csharp
public string Title { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Description**

Gets or sets Description for a Table

```csharp
public string Description { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **AllowOverlap**

Allow table to overlap or not

```csharp
public bool AllowOverlap { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **AllowTextWrap**

Allow text to wrap around table.

```csharp
public bool AllowTextWrap { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **GridColumnWidth**

Sets or gets grid columns width (not really doing anything as far as I can see)

```csharp
public List<int> GridColumnWidth { get; set; }
```

#### Property Value

[List&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **ColumnWidth**

Gets or sets column width for a whole table simplifying setup of column width
 Please note that it assumes first row has the same width as the rest of rows
 which may give false positives if there are multiple values set differently.

```csharp
public List<int> ColumnWidth { get; set; }
```

#### Property Value

[List&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **RowHeight**

Get or Set Table Row Height for 1st row

```csharp
public List<int> RowHeight { get; set; }
```

#### Property Value

[List&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **Cells**

Get all WordTableCells in a table. A short way to loop thru all cells

```csharp
public List<WordTableCell> Cells { get; }
```

#### Property Value

[List&lt;WordTableCell&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **HasNestedTables**

Gets information whether the Table has other nested tables in at least one of the TableCells

```csharp
public bool HasNestedTables { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **NestedTables**

Get all nested tables in the table

```csharp
public List<WordTable> NestedTables { get; }
```

#### Property Value

[List&lt;WordTable&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **IsNestedTable**

Gets information whether the table is nested table (within TableCell)

```csharp
public bool IsNestedTable { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **ParentTable**

Gets nested table parent table if table is nested table

```csharp
public WordTable ParentTable { get; }
```

#### Property Value

[WordTable](./officeimo.word.wordtable.md)<br>

## Constructors

### **WordTable(WordDocument, TableCell, Int32, Int32, WordTableStyle)**

```csharp
public WordTable(WordDocument document, TableCell wordTableCell, int rows, int columns, WordTableStyle tableStyle)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`wordTableCell` TableCell<br>

`rows` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

`columns` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

`tableStyle` [WordTableStyle](./officeimo.word.wordtablestyle.md)<br>

## Methods

### **AddRow(Int32)**

Add row to an existing table with the specified number of columns

```csharp
public void AddRow(int cellsCount)
```

#### Parameters

`cellsCount` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### **AddRow(Int32, Int32)**

Add specified number of rows to an existing table with the specified number of columns

```csharp
public void AddRow(int rowsCount, int cellsCount)
```

#### Parameters

`rowsCount` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

`cellsCount` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### **Remove()**

Remove table from document

```csharp
public void Remove()
```

### **CheckTableProperties()**

Generate table properties for the table if it doesn't exists

```csharp
internal void CheckTableProperties()
```

### **AddComment(String, String, String)**

Add comment to a Table

```csharp
public void AddComment(string author, string initials, string comment)
```

#### Parameters

`author` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`initials` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`comment` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **InsertComment(WordComment, OpenXmlElement, OpenXmlElement, OpenXmlElement)**

```csharp
internal void InsertComment(WordComment wordComment, OpenXmlElement rangeStart, OpenXmlElement rangeEnd, OpenXmlElement reference)
```

#### Parameters

`wordComment` [WordComment](./officeimo.word.wordcomment.md)<br>

`rangeStart` OpenXmlElement<br>

`rangeEnd` OpenXmlElement<br>

`reference` OpenXmlElement<br>
