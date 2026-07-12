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
