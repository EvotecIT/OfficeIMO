# WordTableRow

Namespace: OfficeIMO.Word

```csharp
public class WordTableRow
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordTableRow](./officeimo.word.wordtablerow.md)

## Properties

### **Cells**

Return all cells for given row

```csharp
public List<WordTableCell> Cells { get; }
```

#### Property Value

[List&lt;WordTableCell&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

### **FirstCell**

Return first cell for given row

```csharp
public WordTableCell FirstCell { get; }
```

#### Property Value

[WordTableCell](./officeimo.word.wordtablecell.md)<br>

### **LastCell**

Return last cell for given row

```csharp
public WordTableCell LastCell { get; }
```

#### Property Value

[WordTableCell](./officeimo.word.wordtablecell.md)<br>

### **CellsCount**

Gets cells count

```csharp
public int CellsCount { get; }
```

#### Property Value

[Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### **Height**

Gets or sets the height of a row. The value is stored with
`HeightRuleValues.Exact` to ensure it is preserved even when AutoFit is used.

```csharp
public Nullable<int> Height { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

## Constructors

### **WordTableRow(WordDocument, WordTable)**

```csharp
public WordTableRow(WordDocument document, WordTable wordTable)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`wordTable` [WordTable](./officeimo.word.wordtable.md)<br>

### **WordTableRow(WordTable, TableRow, WordDocument)**

```csharp
public WordTableRow(WordTable wordTable, TableRow row, WordDocument document)
```

#### Parameters

`wordTable` [WordTable](./officeimo.word.wordtable.md)<br>

`row` TableRow<br>

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

## Methods

### **Add(WordTableCell)**

```csharp
public void Add(WordTableCell cell)
```

#### Parameters

`cell` [WordTableCell](./officeimo.word.wordtablecell.md)<br>

### **Remove()**

Remove a row

```csharp
public void Remove()
```

### **MergeVertically(Int32, Int32, Boolean)**

Merges cells starting from the provided column across subsequent rows.

```csharp
public void MergeVertically(int cellIndex, int rowsCount, bool copyParagraphs)
```

#### Parameters

`cellIndex` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>
`rowsCount` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>
`copyParagraphs` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>
