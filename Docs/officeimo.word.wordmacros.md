# WordMacros

Namespace: OfficeIMO.Word

```csharp
public partial class WordDocument
```

## Properties

### **HasMacros**

Indicates whether the document contains a VBA project.

```csharp
public bool HasMacros { get; }
```

### **Macros**

Collection of macro modules contained in the document.

```csharp
public IReadOnlyList<WordMacro> Macros { get; }
```

### **WordMacro**

Represents a single macro module. Use `foreach` over `Macros` and call `macro.Remove()` to delete modules.

## Methods

### **AddMacro(string filePath)**

Adds a VBA project to the document.

```csharp
public void AddMacro(string filePath)
```

#### Parameters

`filePath` String path to the `vbaProject.bin` file.

### **AddMacro(byte[] data)**

Adds a VBA project from a byte array.

```csharp
public void AddMacro(byte[] data)
```

#### Parameters

`data` Byte array containing macro code.

### **ExtractMacros()**

Extracts the VBA project as a byte array.

```csharp
public byte[] ExtractMacros()
```

#### Returns

Byte array with the macro content or `null` when no macros are present.

### **SaveMacros(string filePath)**

Saves the VBA project to a file.

```csharp
public void SaveMacros(string filePath)
```

### **RemoveMacro(string name)**

Removes a single macro module from the document. If the last macro is removed,
the VBA project part is deleted as well.

```csharp
public void RemoveMacro(string name)
```

### **RemoveMacros()**

Removes the VBA project from the document.

```csharp
public void RemoveMacros()
```

## Building a macro

1. Create a macro-enabled document (`.docm`) in Word.
2. Use **Developer → Visual Basic** to write your VBA code.
3. Save the file and close Word.
4. Rename the `.docm` file to `.zip` and extract `vbaProject.bin` from the `word` folder.
5. Use that file with `AddMacro` or `AddMacro(byte[])`.
