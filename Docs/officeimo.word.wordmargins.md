# WordMargins

Namespace: OfficeIMO.Word

```csharp
public class WordMargins
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordMargins](./officeimo.word.wordmargins.md)

## Properties

### **Left**

```csharp
public UInt32Value Left { get; set; }
```

#### Property Value

UInt32Value<br>

### **Right**

```csharp
public UInt32Value Right { get; set; }
```

#### Property Value

UInt32Value<br>

### **Top**

```csharp
public Nullable<int> Top { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Bottom**

```csharp
public Nullable<int> Bottom { get; set; }
```

#### Property Value

[Nullable&lt;Int32&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **HeaderDistance**

```csharp
public UInt32Value HeaderDistance { get; set; }
```

#### Property Value

UInt32Value<br>

### **FooterDistance**

```csharp
public UInt32Value FooterDistance { get; set; }
```

#### Property Value

UInt32Value<br>

### **Gutter**

```csharp
public UInt32Value Gutter { get; set; }
```

#### Property Value

UInt32Value<br>

### **Type**

```csharp
public Nullable<WordMargin> Type { get; set; }
```

#### Property Value

[Nullable&lt;WordMargin&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

## Constructors

### **WordMargins(WordDocument, WordSection)**

```csharp
public WordMargins(WordDocument wordDocument, WordSection wordSection)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`wordSection` [WordSection](./officeimo.word.wordsection.md)<br>

## Methods

### **SetMargins(WordSection, WordMargin)**

```csharp
public static WordSection SetMargins(WordSection wordSection, WordMargin pageMargins)
```

#### Parameters

`wordSection` [WordSection](./officeimo.word.wordsection.md)<br>

`pageMargins` [WordMargin](./officeimo.word.wordmargin.md)<br>

#### Returns

[WordSection](./officeimo.word.wordsection.md)<br>
