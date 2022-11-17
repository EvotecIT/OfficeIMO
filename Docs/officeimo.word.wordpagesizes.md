# WordPageSizes

Namespace: OfficeIMO.Word

```csharp
public class WordPageSizes
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordPageSizes](./officeimo.word.wordpagesizes.md)

## Properties

### **PageSize**

This element specifies the properties (size and orientation) for all pages in the current section.

```csharp
public Nullable<WordPageSize> PageSize { get; set; }
```

#### Property Value

[Nullable&lt;WordPageSize&gt;](https://docs.microsoft.com/en-us/dotnet/api/system.nullable-1)<br>

### **Width**

Get or Set section/page Width

```csharp
public UInt32Value Width { get; set; }
```

#### Property Value

UInt32Value<br>

### **Height**

Get or Set section/page Height

```csharp
public UInt32Value Height { get; set; }
```

#### Property Value

UInt32Value<br>

### **Code**

Get or Set section/page Code

```csharp
public UInt16Value Code { get; set; }
```

#### Property Value

UInt16Value<br>

### **Orientation**

Get or Set section/page Orientation

```csharp
public PageOrientationValues Orientation { get; set; }
```

#### Property Value

PageOrientationValues<br>

### **A3**

```csharp
public static PageSize A3 { get; }
```

#### Property Value

PageSize<br>

### **A4**

```csharp
public static PageSize A4 { get; }
```

#### Property Value

PageSize<br>

### **A5**

```csharp
public static PageSize A5 { get; }
```

#### Property Value

PageSize<br>

### **Executive**

```csharp
public static PageSize Executive { get; }
```

#### Property Value

PageSize<br>

### **A6**

```csharp
public static PageSize A6 { get; }
```

#### Property Value

PageSize<br>

### **B5**

```csharp
public static PageSize B5 { get; }
```

#### Property Value

PageSize<br>

### **Statement**

```csharp
public static PageSize Statement { get; }
```

#### Property Value

PageSize<br>

### **Legal**

```csharp
public static PageSize Legal { get; }
```

#### Property Value

PageSize<br>

### **Letter**

```csharp
public static PageSize Letter { get; }
```

#### Property Value

PageSize<br>

## Constructors

### **WordPageSizes(WordDocument, WordSection)**

Manipulate section/page settings

```csharp
public WordPageSizes(WordDocument wordDocument, WordSection wordSection)
```

#### Parameters

`wordDocument` [WordDocument](./officeimo.word.worddocument.md)<br>

`wordSection` [WordSection](./officeimo.word.wordsection.md)<br>

## Methods

### **GetOrientation(SectionProperties)**

```csharp
internal static PageOrientationValues GetOrientation(SectionProperties sectionProperties)
```

#### Parameters

`sectionProperties` SectionProperties<br>

#### Returns

PageOrientationValues<br>

### **SetOrientation(SectionProperties, PageOrientationValues)**

```csharp
internal static void SetOrientation(SectionProperties sectionProperties, PageOrientationValues pageOrientationValue)
```

#### Parameters

`sectionProperties` SectionProperties<br>

`pageOrientationValue` PageOrientationValues<br>
