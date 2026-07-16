namespace OfficeIMO.Email.Store;

/// <summary>Evidence used to classify a well-known folder role.</summary>
public enum EmailStoreFolderClassificationSource {
    /// <summary>The folder has no well-known classification.</summary>
    None = 0,
    /// <summary>A stable source identifier or format contract identified the folder.</summary>
    SourceIdentifier = 1,
    /// <summary>The source path carried a format-defined folder role.</summary>
    SourcePath = 2,
    /// <summary>The display name supplied a best-effort fallback.</summary>
    DisplayName = 3
}
