# WordHelpers

Namespace: OfficeIMO.Word

```csharp
public class WordHelpers
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordHelpers](./officeimo.word.wordhelpers.md)

## Constructors

### **WordHelpers()**

```csharp
public WordHelpers()
```

## Methods

### **RemoveHeadersAndFooters(String, HeaderFooterValues[])**

Remove selected headers and footers from the document. If no types are supplied all headers and footers are removed.

```csharp
public static void RemoveHeadersAndFooters(string filename, params HeaderFooterValues[] types)
```

#### Parameters

`filename` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>
Document to modify.
`types` [HeaderFooterValues](https://learn.microsoft.com/dotnet/api/documentformat.openxml.wordprocessing.headerfootervalues)[]<br>
Header or footer types to remove.
