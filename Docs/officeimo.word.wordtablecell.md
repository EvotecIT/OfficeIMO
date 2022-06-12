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



```csharp
public Nullable<MergedCellValues> HorizontalMerge { get; set; }
```

#### Property Value

[Nullable&lt;MergedCellValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **VerticalMerge**



```csharp
public Nullable<MergedCellValues> VerticalMerge { get; set; }
```

#### Property Value

[Nullable&lt;MergedCellValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

## Constructors

### **WordTableCell(WordDocument, WordTable, WordTableRow)**



```csharp
public WordTableCell(WordDocument document, WordTable wordTable, WordTableRow wordTableRow)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`wordTable` [WordTable](./officeimo.word.wordtable.md)<br>

`wordTableRow` [WordTableRow](./officeimo.word.wordtablerow.md)<br>

### **WordTableCell(WordDocument, WordTable, WordTableRow, TableCell)**



```csharp
public WordTableCell(WordDocument document, WordTable wordTable, WordTableRow wordTableRow, TableCell tableCell)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`wordTable` [WordTable](./officeimo.word.wordtable.md)<br>

`wordTableRow` [WordTableRow](./officeimo.word.wordtablerow.md)<br>

`tableCell` TableCell<br>

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
