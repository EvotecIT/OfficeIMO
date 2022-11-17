# WordList

Namespace: OfficeIMO.Word

```csharp
public class WordList
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordList](./officeimo.word.wordlist.md)

## Properties

### **IsToc**

This provides a way to set it teams to be treated with heading style during load

```csharp
public bool IsToc { get; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **ListItems**

```csharp
public List<WordParagraph> ListItems { get; }
```

#### Property Value

[List&lt;WordParagraph&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.collections.generic.list-1)<br>

## Constructors

### **WordList(WordDocument, WordSection, Boolean)**

```csharp
public WordList(WordDocument wordDocument, WordSection section, bool isToc)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`section` [WordSection](./officeimo.word.wordsection.md)<br>

`isToc` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **WordList(WordDocument, WordSection, Int32)**

```csharp
public WordList(WordDocument wordDocument, WordSection section, int numberId)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`section` [WordSection](./officeimo.word.wordsection.md)<br>

`numberId` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

## Methods

### **GetNextAbstractNum(Numbering)**

```csharp
internal static int GetNextAbstractNum(Numbering numbering)
```

#### Parameters

`numbering` Numbering<br>

#### Returns

[Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### **GetNextNumberingInstance(Numbering)**

```csharp
internal static int GetNextNumberingInstance(Numbering numbering)
```

#### Parameters

`numbering` Numbering<br>

#### Returns

[Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### **AddList(WordListStyle)**

```csharp
internal void AddList(WordListStyle style)
```

#### Parameters

`style` [WordListStyle](./officeimo.word.wordliststyle.md)<br>

### **AddList(CustomListStyles, String, Int32)**

```csharp
internal void AddList(CustomListStyles style, string levelText, int levelIndex)
```

#### Parameters

`style` [CustomListStyles](./officeimo.word.customliststyles.md)<br>

`levelText` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`levelIndex` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

### **AddItem(String, Int32)**

```csharp
public WordParagraph AddItem(string text, int level)
```

#### Parameters

`text` [String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

`level` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>
