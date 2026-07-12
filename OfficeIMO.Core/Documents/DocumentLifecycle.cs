namespace OfficeIMO.Core;

/// <summary>Controls whether a document may be modified after it is opened.</summary>
public enum DocumentAccessMode {
    /// <summary>The document can be inspected but cannot be changed or saved.</summary>
    ReadOnly,
    /// <summary>The document can be changed and saved explicitly.</summary>
    ReadWrite
}

/// <summary>Controls when changes are written to associated storage.</summary>
public enum DocumentPersistenceMode {
    /// <summary>Changes are written only by an explicit save operation.</summary>
    Explicit,
    /// <summary>Changes are written during disposal and persistence failures are propagated.</summary>
    SaveOnDispose
}

/// <summary>Common lifecycle policy for a newly created document.</summary>
public class DocumentCreateOptions {
    /// <summary>Controls when the document is written to its associated destination.</summary>
    public DocumentPersistenceMode PersistenceMode { get; set; } = DocumentPersistenceMode.Explicit;
}

/// <summary>Common lifecycle policy for a loaded document.</summary>
public class DocumentLoadOptions : DocumentCreateOptions {
    /// <summary>Controls whether the loaded document may be modified.</summary>
    public DocumentAccessMode AccessMode { get; set; } = DocumentAccessMode.ReadWrite;
}
