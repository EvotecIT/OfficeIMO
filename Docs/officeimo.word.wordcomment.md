# WordComment

Namespace: OfficeIMO.Word

```csharp
public class WordComment
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordComment](./officeimo.word.wordcomment.md)

## Properties

### **Id**

ID of a comment

```csharp
public string Id { get; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Text**

Text content of a comment

```csharp
public string Text { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Initials**

Initials of a person who created a comment

```csharp
public string Initials { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Author**

Full name of a person who created a comment

```csharp
public string Author { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **DateTime**

DateTime when the comment was created

```csharp
public DateTime DateTime { get; set; }
```

#### Property Value

[DateTime](https://docs.microsoft.com/en-us/dotnet/api/system.datetime)<br>

## Methods

### **GetNewId(WordDocument)**

```csharp
internal static string GetNewId(WordDocument document)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

#### Returns

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **GetNewId(WordDocument, Comments)**

```csharp
internal static string GetNewId(WordDocument document, Comments comments)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`comments` Comments<br>

#### Returns

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **GetCommentsPart(WordDocument)**

```csharp
internal static Comments GetCommentsPart(WordDocument document)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

#### Returns

Comments<br>

### **Create(WordDocument, String, String, String)**

```csharp
public static WordComment Create(WordDocument document, string author, string initials, string comment)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`author` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`initials` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`comment` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordComment](./officeimo.word.wordcomment.md)<br>

### **GetAllComments(WordDocument)**

```csharp
public static List<WordComment> GetAllComments(WordDocument document)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

#### Returns

[List&lt;WordComment&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>
