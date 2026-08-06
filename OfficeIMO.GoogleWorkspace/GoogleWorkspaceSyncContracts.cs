namespace OfficeIMO.GoogleWorkspace;

/// <summary>Selects how a Google-native document is imported into OfficeIMO.</summary>
public enum GoogleWorkspaceImportMode {
    /// <summary>Export through Google Drive to the corresponding Microsoft Office format.</summary>
    DriveExport,

    /// <summary>Read the native Google Workspace resource and project its supported model.</summary>
    Native
}

/// <summary>Classifies a change found while comparing local and remote document state.</summary>
public enum GoogleWorkspaceDiffKind {
    /// <summary>The local source changed.</summary>
    SourceChange,
    /// <summary>The remote Google Workspace document changed.</summary>
    RemoteChange,
    /// <summary>Both sides changed incompatibly.</summary>
    Conflict,
    /// <summary>Applying the plan requires a lossy action.</summary>
    LossyAction
}
