# WordField

Namespace: OfficeIMO.Word

```csharp
public class WordField
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordField](./officeimo.word.wordfield.md)

## Properties

### **FieldType**

```csharp
public Nullable<WordFieldType> FieldType { get; }
```

#### Property Value

[Nullable&lt;WordFieldType&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **FieldFormat**

```csharp
public Nullable<WordFieldFormat> FieldFormat { get; }
```

#### Property Value

[Nullable&lt;WordFieldFormat&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Field**

```csharp
public string Field { get; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

### **UpdateField**

```csharp
public bool UpdateField { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **LockField**

```csharp
public bool LockField { get; set; }
```

#### Property Value

[Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

### **Text**

```csharp
public string Text { get; set; }
```

#### Property Value

[String](https://docs.microsoft.com/en-us/dotnet/api/system.string)<br>

## Methods

### **AddField(WordParagraph, WordFieldType, Nullable&lt;WordFieldFormat&gt;, Boolean)**

```csharp
public static WordParagraph AddField(WordParagraph paragraph, WordFieldType wordFieldType, Nullable<WordFieldFormat> wordFieldFormat, bool advanced)
```

#### Parameters

`paragraph` [WordParagraph](./officeimo.word.wordparagraph.md)<br>

`wordFieldType` [WordFieldType](./officeimo.word.wordfieldtype.md)<br>

`wordFieldFormat` [Nullable&lt;WordFieldFormat&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

`advanced` [Boolean](https://docs.microsoft.com/en-us/dotnet/api/system.boolean)<br>

#### Returns

[WordParagraph](./officeimo.word.wordparagraph.md)<br>

### **Remove()**

```csharp
public void Remove()
```
