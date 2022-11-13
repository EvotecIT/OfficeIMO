# WordLineChart

Namespace: OfficeIMO.Word

```csharp
public class WordLineChart : WordChart
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordChart](./officeimo.word.wordchart.md) → [WordLineChart](./officeimo.word.wordlinechart.md)

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

### **WordLineChart()**

```csharp
public WordLineChart()
```

## Methods

### **AddLineChart(WordDocument, WordParagraph, Boolean)**

```csharp
public static WordChart AddLineChart(WordDocument wordDocument, WordParagraph paragraph, bool roundedCorners)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`paragraph` [WordParagraph](./officeimo.word.wordparagraph.md)<br>

`roundedCorners` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordChart](./officeimo.word.wordchart.md)<br>

### **CreateLineChart()**

```csharp
internal static LineChart CreateLineChart()
```

#### Returns

LineChart<br>

### **AddLineChartSeries(UInt32Value, String, Color, List&lt;String&gt;, List&lt;Int32&gt;)**

```csharp
internal static LineChartSeries AddLineChartSeries(UInt32Value index, string series, Color color, List<string> categories, List<int> data)
```

#### Parameters

`index` UInt32Value<br>

`series` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`color` Color<br>

`categories` [List&lt;String&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

`data` [List&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

#### Returns

LineChartSeries<br>

### **AddShapeProperties(Color)**

```csharp
internal static ChartShapeProperties AddShapeProperties(Color color)
```

#### Parameters

`color` Color<br>

#### Returns

ChartShapeProperties<br>
