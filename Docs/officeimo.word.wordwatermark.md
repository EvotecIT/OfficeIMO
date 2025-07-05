# WordWatermark

Namespace: OfficeIMO.Word

```csharp
public class WordWatermark
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordWatermark](./officeimo.word.wordwatermark.md)

## Properties

### **Text**

```csharp
public string Text { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

## Constructors

### **WordWatermark(WordDocument, WordSection, WordHeader, WordWatermarkStyle, String, Nullable(Double), Nullable(Double), Double)**

```csharp
public WordWatermark(WordDocument wordDocument, WordSection wordSection, WordHeader wordHeader, WordWatermarkStyle style, string text, double? horizontalOffset = null, double? verticalOffset = null, double scale = 1.0)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`wordSection` [WordSection](./officeimo.word.wordsection.md)<br>

`wordHeader` [WordHeader](./officeimo.word.wordheader.md)<br>

`style` [WordWatermarkStyle](./officeimo.word.wordwatermarkstyle.md)<br>

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
`horizontalOffset` [Nullable(Double)](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>
`verticalOffset` [Nullable(Double)](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>
`scale` [Double](https://docs.microsoft.com/en-us/dotnet/api/system.double)<br>

### **Remove()**

Remove this watermark from the document.

```csharp
public void Remove()
```
