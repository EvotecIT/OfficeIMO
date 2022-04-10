# WordSettings

Namespace: OfficeIMO.Word



```csharp
public class WordSettings
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordSettings](./officeimo.word.wordsettings.md)

## Properties

### **ProtectionType**



```csharp
public Nullable<DocumentProtectionValues> ProtectionType { get; set; }
```

#### Property Value

[Nullable&lt;DocumentProtectionValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ProtectionPassword**



```csharp
public string ProtectionPassword { set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **ZoomPreset**



```csharp
public Nullable<PresetZoomValues> ZoomPreset { get; set; }
```

#### Property Value

[Nullable&lt;PresetZoomValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **ZoomPercentage**



```csharp
public Nullable<int> ZoomPercentage { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **UpdateFieldsOnOpen**



```csharp
public bool UpdateFieldsOnOpen { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **Language**



```csharp
public string Language { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **BackgroundColor**



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
