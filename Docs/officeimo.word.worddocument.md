# WordDocument

Namespace: OfficeIMO.Word

```csharp
public class WordDocument
```

Inheritance [Object](https://docs.microsoft.com/en-us/dotnet/api/system.object) → [WordDocument](./officeimo.word.worddocument.md)

## Methods

### **RemoveSection(Int32)**

Removes the section at the specified index along with its content.
Unused header and footer parts are also cleaned up.

```csharp
public void RemoveSection(int index)
```

#### Parameters

`index` [Int32](https://docs.microsoft.com/en-us/dotnet/api/system.int32)<br>


### **RemoveComment(String)**

Remove comment with the specified id. Alternatively you can call `Remove()` on a `WordComment` instance.

```csharp
public void RemoveComment(string commentId)
```

### **RemoveComment(WordComment)**

Remove the specified comment object.

```csharp
public void RemoveComment(WordComment comment)
```

### **RemoveAllComments()**

Remove all comments from the document.

```csharp
public void RemoveAllComments()
```

### **RemoveWatermark()**

Removes all watermarks from the document including those in headers.
Alternatively remove individual watermarks via the `Watermarks` collection.

```csharp
public void RemoveWatermark()
```

## Properties

### **TrackComments**

Enable or disable tracking of comment changes.

```csharp
public bool TrackComments { get; set; }
```

