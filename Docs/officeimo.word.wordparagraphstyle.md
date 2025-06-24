# WordParagraphStyle

Namespace: OfficeIMO.Word

```csharp
public static class WordParagraphStyle
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordParagraphStyle](./officeimo.word.wordparagraphstyle.md)

## Methods

### **GetStyleDefinition(WordParagraphStyles)**

```csharp
public static Style GetStyleDefinition(WordParagraphStyles style)
```

#### Parameters

`style` [WordParagraphStyles](./officeimo.word.wordparagraphstyles.md)<br>

#### Returns

Style<br>

### **ToStringStyle(WordParagraphStyles)**

```csharp
public static string ToStringStyle(WordParagraphStyles style)
```

#### Parameters

`style` [WordParagraphStyles](./officeimo.word.wordparagraphstyles.md)<br>

#### Returns

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **GetStyle(String)**

```csharp
public static WordParagraphStyles GetStyle(string style)
```

#### Parameters

`style` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordParagraphStyles](./officeimo.word.wordparagraphstyles.md)<br>

### **GetStyle(Int32)**

This method is used to simplify creating TOC List with Headings

```csharp
internal static WordParagraphStyles GetStyle(int level)
```

#### Parameters

`level` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

#### Returns

[WordParagraphStyles](./officeimo.word.wordparagraphstyles.md)<br>

#### Exceptions

[ArgumentOutOfRangeException](https://docs.microsoft.com/en-us/dotnet/api/system.argumentoutofrangeexception)<br>

### **RegisterCustomStyle(String, Style)**

```csharp
public static void RegisterCustomStyle(string styleId, Style styleDefinition)
```

Adds a new style definition keyed by the provided style identifier.

#### Parameters

`styleId` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
`styleDefinition` Style<br>

### **OverrideBuiltInStyle(WordParagraphStyles, Style)**

```csharp
public static void OverrideBuiltInStyle(WordParagraphStyles style, Style styleDefinition)
```

Replaces the definition for a built-in style.
