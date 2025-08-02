# WordPageNumber

Namespace: OfficeIMO.Word

```csharp
public class WordPageNumber
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordPageNumber](./officeimo.word.wordpagenumber.md)

## Properties

### **ParagraphAlignment**

```csharp
public Nullable<JustificationValues> ParagraphAlignment { get; set; }
```

#### Property Value

[Nullable&lt;JustificationValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Paragraph**

```csharp
public WordParagraph Paragraph { get; }
```

#### Property Value

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **Paragraphs**

```csharp
public IReadOnlyList<WordParagraph> Paragraphs { get; }
```

#### Property Value

[IReadOnlyList<WordParagraph>](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlylist-1)<br>

### **Field**

```csharp
public WordField Field { get; }
```

#### Property Value

[WordField](./officeimo.word.wordfield.md)<br>

### **Number**

```csharp
public Nullable<int> Number { get; }
```

#### Property Value

[Nullable<int>](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **CustomFormat**

```csharp
public string CustomFormat { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Example

```csharp
pageNumber.CustomFormat = "00";   // 01, 02, 03...
pageNumber.CustomFormat = "000";  // 001, 002, 003...
pageNumber.CustomFormat = "10-20"; // 10, 11, 12...
```

## Constructors

### **WordPageNumber(WordDocument, WordHeader, WordPageNumberStyle)**

```csharp
public WordPageNumber(WordDocument wordDocument, WordHeader wordHeader, WordPageNumberStyle wordPageNumberStyle)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`wordHeader` [WordHeader](./officeimo.word.wordheader.md)<br>

`wordPageNumberStyle` [WordPageNumberStyle](./officeimo.word.wordpagenumberstyle.md)<br>

### **WordPageNumber(WordDocument, WordFooter, WordPageNumberStyle)**

```csharp
public WordPageNumber(WordDocument wordDocument, WordFooter wordFooter, WordPageNumberStyle wordPageNumberStyle)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`wordFooter` [WordFooter](./officeimo.word.wordfooter.md)<br>

`wordPageNumberStyle` [WordPageNumberStyle](./officeimo.word.wordpagenumberstyle.md)<br>

## Methods

### **AppendText(String)**

```csharp
public WordParagraph AppendText(string text)
```

#### Parameters

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>
