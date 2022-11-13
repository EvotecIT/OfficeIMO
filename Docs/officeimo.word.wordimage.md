# WordImage

Namespace: OfficeIMO.Word

```csharp
public class WordImage
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordImage](./officeimo.word.wordimage.md)

## Properties

### **CompressionQuality**

```csharp
public Nullable<BlipCompressionValues> CompressionQuality { get; set; }
```

#### Property Value

[Nullable&lt;BlipCompressionValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **RelationshipId**

```csharp
public string RelationshipId { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **FilePath**

```csharp
public string FilePath { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **FileName**

```csharp
public string FileName { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Width**

```csharp
public Nullable<double> Width { get; set; }
```

#### Property Value

[Nullable&lt;Double&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Height**

```csharp
public Nullable<double> Height { get; set; }
```

#### Property Value

[Nullable&lt;Double&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **EmuWidth**

```csharp
public Nullable<double> EmuWidth { get; }
```

#### Property Value

[Nullable&lt;Double&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **EmuHeight**

```csharp
public Nullable<double> EmuHeight { get; }
```

#### Property Value

[Nullable&lt;Double&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Shape**

```csharp
public ShapeTypeValues Shape { get; set; }
```

#### Property Value

ShapeTypeValues<br>

### **BlackWiteMode**

```csharp
public BlackWhiteModeValues BlackWiteMode { get; set; }
```

#### Property Value

BlackWhiteModeValues<br>

### **VerticalFlip**

```csharp
public bool VerticalFlip { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **HorizontalFlip**

```csharp
public bool HorizontalFlip { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **Rotation**

```csharp
public int Rotation { get; set; }
```

#### Property Value

[Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### **Wrap**

```csharp
public Nullable<bool> Wrap { get; set; }
```

#### Property Value

[Nullable&lt;Boolean&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

## Constructors

### **WordImage(WordDocument, String, ShapeTypeValues, BlipCompressionValues)**

```csharp
public WordImage(WordDocument document, string filePath, ShapeTypeValues shape, BlipCompressionValues compressionQuality)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`filePath` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`shape` ShapeTypeValues<br>

`compressionQuality` BlipCompressionValues<br>

### **WordImage(WordDocument, Paragraph)**

```csharp
public WordImage(WordDocument document, Paragraph paragraph)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`paragraph` Paragraph<br>

### **WordImage(WordDocument, String, Nullable&lt;Double&gt;, Nullable&lt;Double&gt;, ShapeTypeValues, BlipCompressionValues)**

```csharp
public WordImage(WordDocument document, string filePath, Nullable<double> width, Nullable<double> height, ShapeTypeValues shape, BlipCompressionValues compressionQuality)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`filePath` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`width` [Nullable&lt;Double&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

`height` [Nullable&lt;Double&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

`shape` ShapeTypeValues<br>

`compressionQuality` BlipCompressionValues<br>

### **WordImage(WordDocument, Drawing)**

```csharp
public WordImage(WordDocument document, Drawing drawing)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`drawing` Drawing<br>

## Methods

### **SaveToFile(String)**

Extract image from Word Document and save it to file

```csharp
public void SaveToFile(string fileToSave)
```

#### Parameters

`fileToSave` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Remove()**

```csharp
public void Remove()
```
