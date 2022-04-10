# WordComment

Namespace: OfficeIMO.Word



```csharp
public class WordComment
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordComment](./officeimo.word.wordcomment.md)

## Properties

### **Text**



```csharp
public string Text { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Initials**



```csharp
public string Initials { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Author**



```csharp
public string Author { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **DateTime**



```csharp
public DateTime DateTime { get; set; }
```

#### Property Value

[DateTime](https://docs.microsoft.com/en-us/dotnet/api/system.datetime)<br>

## Methods

### **GetAllComments(WordDocument)**



```csharp
public static List<WordComment> GetAllComments(WordDocument document)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

#### Returns

[List&lt;WordComment&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>
