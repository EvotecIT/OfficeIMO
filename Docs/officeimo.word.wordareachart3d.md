# WordAreaChart3D

Namespace: OfficeIMO.Word

```csharp
public class WordAreaChart3D : WordChart
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordChart](./officeimo.word.wordchart.md) → **WordAreaChart3D**

## Methods

### **CreateArea3DChart(UInt32Value, UInt32Value)**
```csharp
internal static Area3DChart CreateArea3DChart(UInt32Value catAxisId, UInt32Value valAxisId)
```
#### Parameters
`catAxisId` UInt32Value<br>
`valAxisId` UInt32Value<br>
#### Returns
Area3DChart<br>

### **GenerateArea3DChart(Chart)**
```csharp
internal static Chart GenerateArea3DChart(Chart chart)
```
#### Parameters
`chart` Chart<br>
#### Returns
Chart<br>

### **AddArea3DChartSeries(UInt32Value, String, Color, List<string>, List<int>)**
```csharp
internal static AreaChartSeries AddArea3DChartSeries(UInt32Value index, string series, Color color, List<string> categories, List<int> data)
```
#### Parameters
`index` UInt32Value<br>
`series` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
`color` Color<br>
`categories` [List&lt;String&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>
`data` [List&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>
#### Returns
AreaChartSeries<br>

