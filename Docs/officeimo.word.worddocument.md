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

### **HasDocumentVariables**

Indicates if the document contains any document variables.

```csharp
public bool HasDocumentVariables { get; }
```

### **DocumentVariables**

Collection of document variables.

```csharp
public Dictionary<string, string> DocumentVariables { get; }
```

### **GetDocumentVariable(String)**

Return the value of a document variable or <code>null</code> if the variable does not exist.

```csharp
public string GetDocumentVariable(string name)
```

### **SetDocumentVariable(String, String)**

Sets the value of a document variable. Creates it if it does not exist.

```csharp
public void SetDocumentVariable(string name, string value)
```

### **RemoveDocumentVariable(String)**

Remove the document variable with the specified name if present.

```csharp
public void RemoveDocumentVariable(string name)
```

### **RemoveDocumentVariableAt(Int32)**

Remove the document variable at the given index.

```csharp
public void RemoveDocumentVariableAt(int index)
```

### **GetDocumentVariables()**

Returns a read-only collection of all document variables.

```csharp
public IReadOnlyDictionary<string, string> GetDocumentVariables()
```

