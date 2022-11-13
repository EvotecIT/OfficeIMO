# WordTableStyles

Namespace: OfficeIMO.Word

```csharp
public static class WordTableStyles
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordTableStyles](./officeimo.word.wordtablestyles.md)

## Methods

### **GetStyle(String)**

```csharp
public static WordTableStyle GetStyle(string style)
```

#### Parameters

`style` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordTableStyle](./officeimo.word.wordtablestyle.md)<br>

### **GetStyle(WordTableStyle)**

```csharp
public static TableStyle GetStyle(WordTableStyle style)
```

#### Parameters

`style` [WordTableStyle](./officeimo.word.wordtablestyle.md)<br>

#### Returns

TableStyle<br>

### **IsAvailableStyle(Styles, WordTableStyle)**

Verifies whether table style is available in document or not

```csharp
internal static bool IsAvailableStyle(Styles styles, WordTableStyle style)
```

#### Parameters

`styles` Styles<br>

`style` [WordTableStyle](./officeimo.word.wordtablestyle.md)<br>

#### Returns

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **GetStyleDefinition(WordTableStyle)**

```csharp
public static Style GetStyleDefinition(WordTableStyle style)
```

#### Parameters

`style` [WordTableStyle](./officeimo.word.wordtablestyle.md)<br>

#### Returns

Style<br>
