namespace OfficeIMO.IWork;

/// <summary>Identifies the iWork application that owns a source package.</summary>
public enum IWorkDocumentKind {
    /// <summary>Apple Pages word-processing document.</summary>
    Pages,
    /// <summary>Apple Numbers spreadsheet.</summary>
    Numbers,
    /// <summary>Apple Keynote presentation.</summary>
    Keynote
}

/// <summary>Identifies the physical layout used by an iWork source.</summary>
public enum IWorkContainerKind {
    /// <summary>A ZIP archive containing the iWork package.</summary>
    ZipPackage,
    /// <summary>An unpacked iWork package directory.</summary>
    DirectoryBundle,
    /// <summary>A ZIP archive whose primary index is nested in an inner package.</summary>
    ZipPackageWithNestedIndex
}

/// <summary>Controls how a destination adapter should convert an opened iWork source.</summary>
public enum IWorkConversionMode {
    /// <summary>Prefer editable semantic content and use an embedded raster preview only when no supported structure is available.</summary>
    Auto,
    /// <summary>Require editable semantic reconstruction and fail when supported structure cannot be recovered.</summary>
    EditableOnly,
    /// <summary>Use an embedded visual preview without claiming editable reconstruction.</summary>
    VisualOnly
}

/// <summary>Describes the representation produced by an iWork conversion.</summary>
public enum IWorkProjectionKind {
    /// <summary>Editable content reconstructed from supported source structures.</summary>
    EditableReconstruction,
    /// <summary>Visual output based on an embedded preview because editable reconstruction was unavailable.</summary>
    VisualFallback
}

/// <summary>Describes how much of a source an embedded preview is known to cover.</summary>
public enum IWorkVisualCoverage {
    /// <summary>The preview's coverage cannot be established.</summary>
    Unknown,
    /// <summary>The preview represents only the first page or a composite thumbnail.</summary>
    FirstPageOrCompositePreview,
    /// <summary>The preview represents the complete document.</summary>
    FullDocument
}

/// <summary>Severity of an iWork read or projection diagnostic.</summary>
public enum IWorkDiagnosticSeverity {
    /// <summary>Informational recovery detail.</summary>
    Information,
    /// <summary>Recoverable limitation or possible fidelity loss.</summary>
    Warning,
    /// <summary>Failure that prevents the requested projection.</summary>
    Error
}

/// <summary>Typed value recovered from a Numbers table cell.</summary>
public enum IWorkCellKind {
    /// <summary>An empty cell.</summary>
    Empty,
    /// <summary>A text value.</summary>
    Text,
    /// <summary>A numeric value.</summary>
    Number,
    /// <summary>A Boolean value.</summary>
    Boolean,
    /// <summary>A date or date-time value.</summary>
    DateTime,
    /// <summary>A duration value.</summary>
    Duration,
    /// <summary>A formula expression or calculated value.</summary>
    Formula,
    /// <summary>A source cell error.</summary>
    Error
}

/// <summary>Identifies one drawable recovered from a Keynote slide.</summary>
public enum IWorkKeynoteDrawableKind {
    /// <summary>A positioned rich-text shape.</summary>
    TextBox,
    /// <summary>An embedded raster image.</summary>
    Image,
    /// <summary>An editable table.</summary>
    Table
}

/// <summary>Identifies one drawable recovered from a Pages document.</summary>
public enum IWorkPagesDrawableKind {
    /// <summary>A positioned rich-text shape.</summary>
    TextBox,
    /// <summary>An embedded raster image.</summary>
    Image,
    /// <summary>An editable table.</summary>
    Table
}
