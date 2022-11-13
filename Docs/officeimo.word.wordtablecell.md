# WordTableCell

Namespace: OfficeIMO.Word

```csharp
public class WordTableCell
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordTableCell](./officeimo.word.wordtablecell.md)

## Fields

### **Borders**

```csharp
public WordTableCellBorder Borders;
```

## Properties

### **Paragraphs**

```csharp
public List<WordParagraph> Paragraphs { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **HorizontalMerge**

Gets or Sets Horizontal Merge for a Table Cell

```csharp
public Nullable<MergedCellValues> HorizontalMerge { get; set; }
```

#### Property Value

[Nullable&lt;MergedCellValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **VerticalMerge**

Gets or Sets Vertical Merge for a Table Cell

```csharp
public Nullable<MergedCellValues> VerticalMerge { get; set; }
```

#### Property Value

[Nullable&lt;MergedCellValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ShadingFillColorHex**

Get or set the background color of the cell using hexadecimal color code.

```csharp
public string ShadingFillColorHex { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **ShadingPattern**

Get or set the background pattern of a cell

```csharp
public Nullable<ShadingPatternValues> ShadingPattern { get; set; }
```

#### Property Value

[Nullable&lt;ShadingPatternValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ShadingFillColor**

Get or set the background color of a cell using SixLabors.Color

```csharp
public Nullable<Color> ShadingFillColor { get; set; }
```

#### Property Value

[Nullable&lt;Color&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Width**

Gets or sets cell width

```csharp
public Nullable<int> Width { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **TextDirection**

Gets or sets text direction in a Table Cell

```csharp
public Nullable<TextDirectionValues> TextDirection { get; set; }
```

#### Property Value

[Nullable&lt;TextDirectionValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

## Constructors

### **WordTableCell(WordDocument, WordTable, WordTableRow)**

Create a WordTableCell and add it to given Table Row

```csharp
public WordTableCell(WordDocument document, WordTable wordTable, WordTableRow wordTableRow)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`wordTable` [WordTable](./officeimo.word.wordtable.md)<br>

`wordTableRow` [WordTableRow](./officeimo.word.wordtablerow.md)<br>

## Methods

### **Remove()**

Remove a cell from a table

```csharp
public void Remove()
```

### **MergeHorizontally(Int32, Boolean)**

Merges two or more cells together horizontally.
 Provides ability to move or delete content of merged cells into single cell

```csharp
public void MergeHorizontally(int cellsCount, bool copyParagraphs)
```

#### Parameters

`cellsCount` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

`copyParagraphs` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **SplitHorizontally(Int32)**

Splits (unmerge) cells that were merged

```csharp
public void SplitHorizontally(int cellsCount)
```

#### Parameters

`cellsCount` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### **MergeVertically(Int32, Boolean)**

Merges two or more cells together vertically

```csharp
public void MergeVertically(int cellsCount, bool copyParagraphs)
```

#### Parameters

`cellsCount` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

`copyParagraphs` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **AddTable(Int32, Int32, WordTableStyle, Boolean)**

Add table to a table cell (nested table)

```csharp
public WordTable AddTable(int rows, int columns, WordTableStyle tableStyle, bool removePrecedingParagraph)
```

#### Parameters

`rows` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

`columns` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

`tableStyle` [WordTableStyle](./officeimo.word.wordtablestyle.md)<br>

`removePrecedingParagraph` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordTable](./officeimo.word.wordtable.md)<br>
