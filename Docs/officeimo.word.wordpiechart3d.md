# WordPieChart3D

Namespace: OfficeIMO.Word

```csharp
public class WordPieChart3D : WordChart
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordChart](./officeimo.word.wordchart.md) → [WordPieChart3D](./officeimo.word.wordpiechart3d.md)

## Methods

### **CreatePie3DChart(Chart)**
```csharp
internal static Chart CreatePie3DChart(Chart chart)
```

#### Parameters
`chart` Chart<br>

#### Returns
Chart<br>

### **GeneratePie3DChart(Chart)**
```csharp
internal static Chart GeneratePie3DChart(Chart chart)
```

#### Parameters
`chart` Chart<br>

#### Returns
Chart<br>

### **AddPie3DChartSeries(UInt32Value, String, Color, List<string>, List<int>)**
```csharp
internal static PieChartSeries AddPie3DChartSeries(UInt32Value index, string series, Color color, List<string> categories, List<int> data)
```

#### Parameters
`index` UInt32Value<br>
`series` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
`color` Color<br>
`categories` [List&lt;String&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>
`data` [List&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

#### Returns
PieChartSeries<br>

