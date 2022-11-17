# WordSettings

Namespace: OfficeIMO.Word

```csharp
public class WordSettings
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordSettings](./officeimo.word.wordsettings.md)

## Properties

### **ProtectionType**

Get or set Protection Type for the document

```csharp
public Nullable<DocumentProtectionValues> ProtectionType { get; set; }
```

#### Property Value

[Nullable&lt;DocumentProtectionValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ProtectionPassword**

Set a Protection Password for the document

```csharp
public string ProtectionPassword { set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **ZoomPreset**

Get or set Zoom Preset for the document

```csharp
public Nullable<PresetZoomValues> ZoomPreset { get; set; }
```

#### Property Value

[Nullable&lt;PresetZoomValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ZoomPercentage**

Get or Set Zoome Percentage for the document

```csharp
public Nullable<int> ZoomPercentage { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **UpdateFieldsOnOpen**

Tell Word to update fields when opening word.
 Without this option the document fields won't be refreshed until manual intervention.

```csharp
public bool UpdateFieldsOnOpen { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **FontSize**

Gets or Sets default font size for the whole document. Default is 11.

```csharp
public Nullable<int> FontSize { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **FontSizeComplexScript**

Gets or Sets default font size complex script for the whole document. Default is 11.

```csharp
public Nullable<int> FontSizeComplexScript { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **FontFamily**

Gets or Sets default font family for the whole document.

```csharp
public string FontFamily { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Language**

Gets or Sets default language for the whole document. Default is en-Us.

```csharp
public string Language { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **BackgroundColor**

Gets or Sets default Background Color for the whole document

```csharp
public string BackgroundColor { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

## Constructors

### **WordSettings(WordDocument)**

```csharp
public WordSettings(WordDocument document)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

## Methods

### **RemoveProtection()**

Remove protection from document (if it's set).

```csharp
public void RemoveProtection()
```

### **SetBackgroundColor(String)**

```csharp
public WordSettings SetBackgroundColor(string backgroundColor)
```

#### Parameters

`backgroundColor` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordSettings](./officeimo.word.wordsettings.md)<br>

### **SetBackgroundColor(Color)**

```csharp
public WordSettings SetBackgroundColor(Color backgroundColor)
```

#### Parameters

`backgroundColor` Color<br>

#### Returns

[WordSettings](./officeimo.word.wordsettings.md)<br>
