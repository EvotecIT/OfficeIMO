# WordBreak

Namespace: OfficeIMO.Word

Represents a break in the text.
 Be it page break, soft break, column or text wrapping

```csharp
public class WordBreak
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordBreak](./officeimo.word.wordbreak.md)

## Properties

### **BreakType**

Get type of Break in given paragraph

```csharp
public Nullable<BreakValues> BreakType { get; }
```

#### Property Value

[Nullable&lt;BreakValues&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

## Constructors

### **WordBreak(WordDocument, Paragraph, Run)**

Create new instance of WordBreak

```csharp
public WordBreak(WordDocument document, Paragraph paragraph, Run run)
```

#### Parameters

`document` [WordDocument](./officeimo.word.worddocument.md)<br>

`paragraph` Paragraph<br>

`run` Run<br>

## Methods

### **Remove(Boolean)**

Remove the break from WordDocument. By default it removes break without removing paragraph.
 If you want paragraph removed please use IncludingParagraph bool.
 Please remember a paragraph can hold multiple other elements.

```csharp
public void Remove(bool includingParagraph)
```

#### Parameters

`includingParagraph` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>
