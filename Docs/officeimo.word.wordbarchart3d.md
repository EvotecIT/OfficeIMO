# WordBarChart3D

Namespace: OfficeIMO.Word

```csharp
public class WordBarChart3D : WordChart
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordChart](./officeimo.word.wordchart.md) → [WordBarChart3D](./officeimo.word.wordbarchart3d.md)

## Properties

### **BarGrouping**

```csharp
public Nullable<BarGroupingValues> BarGrouping { get; set; }
```

#### Property Value

[Nullable&lt;BarGroupingValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **BarDirection**

```csharp
public Nullable<BarDirectionValues> BarDirection { get; set; }
```

#### Property Value

[Nullable&lt;BarDirectionValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **RoundedCorners**

```csharp
public bool RoundedCorners { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **Categories**

```csharp
public List<string> Categories { get; set; }
```

#### Property Value

[List&lt;String&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

## Constructors

### **WordBarChart3D()**

```csharp
public WordBarChart3D()
```

## Methods

### **AddBarChart3D(WordDocument, WordParagraph)**

```csharp
public static WordBarChart3D AddBarChart3D(WordDocument wordDocument, WordParagraph paragraph)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`paragraph` [WordParagraph](./officeimo.word.wordparagraph.md)<br>

#### Returns

[WordBarChart3D](./officeimo.word.wordbarchart3d.md)<br>

### **GenerateRoundedCorners()**

```csharp
public RoundedCorners GenerateRoundedCorners()
```

#### Returns

RoundedCorners<br>

### **GenerateChart()**

```csharp
public Chart GenerateChart()
```

#### Returns

Chart<br>

### **GenerateRun()**

```csharp
public Run GenerateRun()
```

#### Returns

Run<br>
