# WordBackground

Namespace: OfficeIMO.Word

```csharp
public class WordBackground
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordBackground](./officeimo.word.wordbackground.md)

## Properties

### **Color**

```csharp
public string Color { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

## Constructors

### **WordBackground(WordDocument)**

```csharp
public WordBackground(WordDocument document)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

### **WordBackground(WordDocument, Color)**

```csharp
public WordBackground(WordDocument document, Color color)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`color` Color<br>

## Methods

### **SetColorHex(String)**

```csharp
public WordBackground SetColorHex(string color)
```

#### Parameters

`color` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordBackground](./officeimo.word.wordbackground.md)<br>

### **SetColor(Color)**

```csharp
public WordBackground SetColor(Color color)
```

#### Parameters

`color` Color<br>

#### Returns

[WordBackground](./officeimo.word.wordbackground.md)<br>

### **SetImage(String, Double?, Double?)**

```csharp
public WordBackground SetImage(string filePath, double? width = null, double? height = null)
```

#### Parameters

`filePath` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
`width` [Double](https://docs.microsoft.com/en-us/dotnet/api/system.double)?<br>
`height` [Double](https://docs.microsoft.com/en-us/dotnet/api/system.double)?<br>

#### Returns

[WordBackground](./officeimo.word.wordbackground.md)<br>

### **SetImage(Stream, String, Double?, Double?)**

```csharp
public WordBackground SetImage(Stream imageStream, string fileName, double? width = null, double? height = null)
```

#### Parameters

`imageStream` [Stream](https://docs.microsoft.com/en-us/dotnet/api/system.io.stream)<br>
`fileName` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
`width` [Double](https://docs.microsoft.com/en-us/dotnet/api/system.double)?<br>
`height` [Double](https://docs.microsoft.com/en-us/dotnet/api/system.double)?<br>

#### Returns

[WordBackground](./officeimo.word.wordbackground.md)<br>
