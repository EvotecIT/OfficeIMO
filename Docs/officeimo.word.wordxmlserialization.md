# WordXmlSerialization

Namespace: OfficeIMO.Word

```csharp
public static class WordXmlExtensions
```

Provides helper methods to convert Word elements to and from raw XML.

## Methods

### **ToXml(WordParagraph)**

```csharp
public static string ToXml(this WordParagraph paragraph)
```

Returns the `OuterXml` of the underlying `Paragraph` element.

### **AddParagraphFromXml(WordDocument, string)**

```csharp
public static WordParagraph AddParagraphFromXml(this WordDocument document, string xml)
```

Creates a paragraph from an XML string and appends it to the specified document.

## Example

```csharp
using OfficeIMO.Word;

using var doc = WordDocument.Create("example.docx");
var p = doc.AddParagraph("Hello world");
string xml = p.ToXml();

// insert a copy of the paragraph
var copy = doc.AddParagraphFromXml(xml);

doc.Save(true);
```
