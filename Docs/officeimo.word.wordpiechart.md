# WordPieChart

Namespace: OfficeIMO.Word

```csharp
public class WordPieChart : WordChart
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordChart](./officeimo.word.wordchart.md) → [WordPieChart](./officeimo.word.wordpiechart.md)

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

### **WordPieChart()**

```csharp
public WordPieChart()
```

## Methods

### **AddPieChart(WordDocument, WordParagraph, Boolean)**

```csharp
public static WordChart AddPieChart(WordDocument wordDocument, WordParagraph paragraph, bool roundedCorners)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`paragraph` [WordParagraph](./officeimo.word.wordparagraph.md)<br>

`roundedCorners` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordChart](./officeimo.word.wordchart.md)<br>

### **CreatePieChart(Chart)**

```csharp
internal static Chart CreatePieChart(Chart chart)
```

#### Parameters

`chart` Chart<br>

#### Returns

Chart<br>

### **AddPieChartSeries(UInt32Value, String, Color, List&lt;String&gt;, List&lt;Int32&gt;)**

```csharp
internal static PieChartSeries AddPieChartSeries(UInt32Value index, string series, Color color, List<string> categories, List<int> data)
```

#### Parameters

`index` UInt32Value<br>

`series` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`color` Color<br>

`categories` [List&lt;String&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

`data` [List&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

#### Returns

PieChartSeries<br>
