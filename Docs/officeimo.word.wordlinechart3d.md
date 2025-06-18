# WordLineChart3D

Namespace: OfficeIMO.Word

```csharp
public class WordLineChart3D : WordChart
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordChart](./officeimo.word.wordchart.md) → [WordLineChart3D](./officeimo.word.wordlinechart3d.md)

## Methods

### **CreateLine3DChart(UInt32Value, UInt32Value)**
```csharp
internal static Line3DChart CreateLine3DChart(UInt32Value catAxisId, UInt32Value valAxisId)
```

#### Parameters
`catAxisId` UInt32Value<br>
`valAxisId` UInt32Value<br>

#### Returns
Line3DChart<br>

### **GenerateLine3DChart(Chart)**
```csharp
internal static Chart GenerateLine3DChart(Chart chart)
```

#### Parameters
`chart` Chart<br>

#### Returns
Chart<br>

### **AddLine3DChartSeries(UInt32Value, String, Color, List<string>, List<int>)**
```csharp
internal static LineChartSeries AddLine3DChartSeries(UInt32Value index, string series, Color color, List<string> categories, List<int> data)
```

#### Parameters
`index` UInt32Value<br>
`series` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
`color` Color<br>
`categories` [List&lt;String&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>
`data` [List&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

#### Returns
LineChartSeries<br>
