# WordTable

Namespace: OfficeIMO.Word



```csharp
public class WordTable
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordTable](./officeimo.word.wordtable.md)

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

### **FirstRow**

Specifies that the first row conditional formatting shall be applied to the table.

```csharp
public Nullable<bool> FirstRow { get; set; }
```

#### Property Value

[Nullable&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **LastRow**

Specifies that the last row conditional formatting shall be applied to the table.

```csharp
public Nullable<bool> LastRow { get; set; }
```

#### Property Value

[Nullable&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **FirstColumn**

Specifies that the first column conditional formatting shall be applied to the table.

```csharp
public Nullable<bool> FirstColumn { get; set; }
```

#### Property Value

[Nullable&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **LastColumn**

Specifies that the last column conditional formatting shall be applied to the table.

```csharp
public Nullable<bool> LastColumn { get; set; }
```

#### Property Value

[Nullable&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **NoHorizontalBand**

Specifies that the horizontal banding conditional formatting shall not be applied to the table.

```csharp
public Nullable<bool> NoHorizontalBand { get; set; }
```

#### Property Value

[Nullable&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **NoVerticalBand**

Specifies that the vertical banding conditional formatting shall not be applied to the table.

```csharp
public Nullable<bool> NoVerticalBand { get; set; }
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

## Methods

### **AddRow(Int32)**



```csharp
public void AddRow(int cellsCount)
```

#### Parameters

`cellsCount` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### **AddRow(Int32, Int32)**



```csharp
public void AddRow(int rowsCount, int cellsCount)
```

#### Parameters

`rowsCount` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

`cellsCount` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### **Remove()**



```csharp
public void Remove()
```
