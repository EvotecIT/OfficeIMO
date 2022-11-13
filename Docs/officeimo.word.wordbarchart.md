# WordBarChart

Namespace: OfficeIMO.Word

```csharp
public class WordBarChart : WordChart
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordChart](./officeimo.word.wordchart.md) → [WordBarChart](./officeimo.word.wordbarchart.md)

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

### **WordBarChart()**

```csharp
public WordBarChart()
```

## Methods

### **AddBarChart(WordDocument, WordParagraph, Boolean)**

```csharp
public static WordChart AddBarChart(WordDocument wordDocument, WordParagraph paragraph, bool roundedCorners)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`paragraph` [WordParagraph](./officeimo.word.wordparagraph.md)<br>

`roundedCorners` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordChart](./officeimo.word.wordchart.md)<br>

### **CreateBarChart(BarDirectionValues)**

```csharp
internal static BarChart CreateBarChart(BarDirectionValues barDirection)
```

#### Parameters

`barDirection` BarDirectionValues<br>

#### Returns

BarChart<br>

### **AddShapeProperties(Color)**

```csharp
internal static ChartShapeProperties AddShapeProperties(Color color)
```

#### Parameters

`color` Color<br>

#### Returns

ChartShapeProperties<br>

### **AddBarChartSeries(UInt32Value, String, Color, List&lt;String&gt;, List&lt;Int32&gt;)**

```csharp
internal static BarChartSeries AddBarChartSeries(UInt32Value index, string series, Color color, List<string> categories, List<int> data)
```

#### Parameters

`index` UInt32Value<br>

`series` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`color` Color<br>

`categories` [List&lt;String&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

`data` [List&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

#### Returns

BarChartSeries<br>
