# WordHyperLink

Namespace: OfficeIMO.Word

```csharp
public class WordHyperLink
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordHyperLink](./officeimo.word.wordhyperlink.md)

## Properties

### **Uri**

```csharp
public Uri Uri { get; set; }
```

#### Property Value

Uri<br>

### **Id**

Specifies a location in the target of the hyperlink, in the case in which the link is an external link.

```csharp
public string Id { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **IsEmail**

```csharp
public bool IsEmail { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **EmailAddress**

```csharp
public string EmailAddress { get; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **History**

```csharp
public bool History { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **DocLocation**

Specifies a location in the target of the hyperlink, in the case in which the link is an external link.

```csharp
public string DocLocation { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Anchor**

Specifies the name of a bookmark within the document.
 See Bookmark. If the attribute is omitted, then the default behavior is to navigate to the start of the document.
 If the r:id attribute is specified, then the anchor attribute is ignored.

```csharp
public string Anchor { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Tooltip**

```csharp
public string Tooltip { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **TargetFrame**

```csharp
public Nullable<TargetFrame> TargetFrame { get; set; }
```

#### Property Value

[Nullable&lt;TargetFrame&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **IsHttp**

```csharp
public bool IsHttp { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **Scheme**

```csharp
public string Scheme { get; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **Text**

```csharp
public string Text { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

## Constructors

### **WordHyperLink(WordDocument, Paragraph, Hyperlink)**

```csharp
public WordHyperLink(WordDocument document, Paragraph paragraph, Hyperlink hyperlink)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`paragraph` Paragraph<br>

`hyperlink` Hyperlink<br>

## Methods

### **Remove(Boolean)**

Removes hyperlink. When specified to remove paragraph it will only do so,
 if paragraph is empty or contains only paragraph properties.

```csharp
public void Remove(bool includingParagraph)
```

#### Parameters

`includingParagraph` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **AddHyperLink(WordParagraph, String, String, Boolean, String, Boolean)**

```csharp
public static WordParagraph AddHyperLink(WordParagraph paragraph, string text, string anchor, bool addStyle, string tooltip, bool history)
```

#### Parameters

`paragraph` [WordParagraph](./officeimo.word.wordparagraph.md)<br>

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`anchor` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`addStyle` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

`tooltip` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`history` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **AddHyperLink(WordParagraph, String, Uri, Boolean, String, Boolean)**

```csharp
public static WordParagraph AddHyperLink(WordParagraph paragraph, string text, Uri uri, bool addStyle, string tooltip, bool history)
```

#### Parameters

`paragraph` [WordParagraph](./officeimo.word.wordparagraph.md)<br>

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`uri` Uri<br>

`addStyle` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

`tooltip` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`history` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>
