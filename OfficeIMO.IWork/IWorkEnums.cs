namespace OfficeIMO.IWork;

/// <summary>Identifies the iWork application that owns a source package.</summary>
public enum IWorkDocumentKind {
    Pages,
    Numbers,
    Keynote
}

/// <summary>Identifies the physical layout used by an iWork source.</summary>
public enum IWorkContainerKind {
    ZipPackage,
    DirectoryBundle,
    ZipPackageWithNestedIndex
}

/// <summary>Controls how a semantic owner should project an iWork source.</summary>
public enum IWorkImportMode {
    /// <summary>Prefer editable semantic content and use an embedded raster preview only when no supported structure is available.</summary>
    Auto,
    /// <summary>Require editable semantic reconstruction and fail when supported structure cannot be recovered.</summary>
    EditableOnly,
    /// <summary>Use an embedded visual preview without claiming editable reconstruction.</summary>
    VisualOnly
}

/// <summary>Describes the representation produced by an iWork import.</summary>
public enum IWorkProjectionKind {
    EditableReconstruction,
    VisualFallback
}

/// <summary>Describes how much of a source an embedded preview is known to cover.</summary>
public enum IWorkVisualCoverage {
    Unknown,
    FirstPageOrCompositePreview,
    FullDocument
}

/// <summary>Severity of an iWork read or projection diagnostic.</summary>
public enum IWorkDiagnosticSeverity {
    Information,
    Warning,
    Error
}

/// <summary>Typed value recovered from a Numbers table cell.</summary>
public enum IWorkCellKind {
    Empty,
    Text,
    Number,
    Boolean,
    DateTime,
    Duration,
    Formula,
    Error
}
