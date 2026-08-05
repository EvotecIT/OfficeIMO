using System;

namespace OfficeIMO.Drawing;

/// <summary>Identifies the Office document family participating in a conversion.</summary>
public enum OfficeDocumentFamily {
    /// <summary>Word-processing documents.</summary>
    Word,
    /// <summary>Spreadsheet workbooks.</summary>
    Excel,
    /// <summary>Presentations, templates, and slide shows.</summary>
    PowerPoint
}

/// <summary>Identifies the logical kind of Office artifact.</summary>
public enum OfficeDocumentKind {
    /// <summary>A normal document, workbook, or presentation.</summary>
    Document,
    /// <summary>A reusable template.</summary>
    Template,
    /// <summary>A PowerPoint slide show.</summary>
    SlideShow,
    /// <summary>An Office add-in.</summary>
    AddIn,
    /// <summary>A legacy Excel workspace.</summary>
    Workspace
}

/// <summary>Identifies the generation of an Office file format.</summary>
public enum OfficeFormatGeneration {
    /// <summary>A legacy Office 97-2003 binary format.</summary>
    Legacy,
    /// <summary>A modern Office format.</summary>
    Modern
}

/// <summary>Identifies the physical encoding used by an Office file format.</summary>
public enum OfficeFormatEncoding {
    /// <summary>OLE Compound File Binary container and family-specific binary records.</summary>
    CompoundBinary,
    /// <summary>Open Packaging Conventions package containing XML parts.</summary>
    OpenXml,
    /// <summary>Open Packaging Conventions package containing family-specific binary parts.</summary>
    BinaryOpenXml
}

/// <summary>Controls how a conversion balances editability, appearance, and preservation.</summary>
public enum OfficeCompatibilityMode {
    /// <summary>Accept only native or semantically equivalent target representations.</summary>
    StrictNative,
    /// <summary>Prefer an editable approximation over a static visual representation.</summary>
    PreferEditable,
    /// <summary>Prefer visual fidelity when an editable representation is unavailable.</summary>
    PreferVisual,
    /// <summary>Use the safest available native, editable, visual, or embedded-source fallback.</summary>
    BestEffort,
    /// <summary>Permit opaque retention for round trips without claiming editable conversion parity.</summary>
    PreservationOnly
}

/// <summary>Describes how one source feature is represented by a conversion.</summary>
public enum OfficeCompatibilityState {
    /// <summary>The target uses its native editable representation.</summary>
    Native,
    /// <summary>The target uses a semantically equivalent editable representation.</summary>
    Equivalent,
    /// <summary>The target uses an editable but incomplete approximation.</summary>
    Approximated,
    /// <summary>The feature is replaced with a static visual representation.</summary>
    Rasterized,
    /// <summary>The original source payload is embedded alongside a fallback representation.</summary>
    EmbeddedSource,
    /// <summary>The feature is retained opaquely and is not projected to the editable model.</summary>
    PreservedOpaque,
    /// <summary>The feature is deliberately omitted under the selected policy.</summary>
    Dropped,
    /// <summary>The conversion is refused because no accepted safe representation exists.</summary>
    Blocked
}

/// <summary>Identifies the severity of a compatibility finding.</summary>
public enum OfficeCompatibilitySeverity {
    /// <summary>Informational compatibility detail.</summary>
    Information,
    /// <summary>A finding requiring review.</summary>
    Warning,
    /// <summary>A finding that prevented conversion.</summary>
    Error
}

/// <summary>Identifies fidelity dimensions affected by a compatibility decision.</summary>
[Flags]
public enum OfficeCompatibilityImpact {
    /// <summary>No fidelity dimension is affected.</summary>
    None = 0,
    /// <summary>Document meaning or structure changes.</summary>
    Semantic = 1,
    /// <summary>Rendered appearance changes.</summary>
    Visual = 2,
    /// <summary>Formula, macro, action, animation, or other behavior changes.</summary>
    Behavioral = 4,
    /// <summary>The target is less editable than the source.</summary>
    Editability = 8,
    /// <summary>Opaque source records or payloads are not retained.</summary>
    Carrier = 16,
    /// <summary>The conversion changes a security property or active-content boundary.</summary>
    Security = 32
}

/// <summary>Identifies a direction in a legacy/modern Office compatibility contract.</summary>
public enum OfficeCapabilityLane {
    /// <summary>Legacy input projected into the normal editable OfficeIMO model.</summary>
    LegacyImport,
    /// <summary>A new legacy artifact authored from the normal OfficeIMO model.</summary>
    NewLegacyWrite,
    /// <summary>An imported legacy artifact edited and saved back to its legacy format.</summary>
    LegacyRoundTrip,
    /// <summary>A modern artifact converted to a legacy format.</summary>
    ModernToLegacy,
    /// <summary>A legacy artifact converted to a modern format.</summary>
    LegacyToModern
}

/// <summary>Describes what the legacy format itself can carry for a feature.</summary>
public enum OfficeCapabilityRepresentability {
    /// <summary>The legacy format has a native editable representation.</summary>
    Native,
    /// <summary>The legacy format has a different but semantically equivalent representation.</summary>
    Equivalent,
    /// <summary>The legacy format can carry a documented approximation.</summary>
    Approximation,
    /// <summary>The feature can only be retained as an opaque carrier.</summary>
    Opaque,
    /// <summary>The legacy format has no safe representation for the feature.</summary>
    NotRepresentable
}

/// <summary>Describes current implementation coverage for a capability lane.</summary>
public enum OfficeCapabilityCoverageState {
    /// <summary>The lane does not apply to this feature or source format.</summary>
    NotApplicable,
    /// <summary>The feature is projected or written through a native editable representation.</summary>
    Native,
    /// <summary>The feature is projected or written through a semantically equivalent editable representation.</summary>
    Equivalent,
    /// <summary>The feature is projected or written through a documented editable approximation.</summary>
    Approximated,
    /// <summary>The feature is deliberately flattened to a static visual representation.</summary>
    Rasterized,
    /// <summary>The original source payload is embedded alongside an editable or visual fallback.</summary>
    EmbeddedSource,
    /// <summary>The original carrier is retained without editable projection.</summary>
    PreservedOpaque,
    /// <summary>The feature is deliberately omitted and the loss is reported.</summary>
    Dropped,
    /// <summary>The operation is deliberately refused to prevent unaccepted loss.</summary>
    Blocked,
    /// <summary>The format can represent the feature, but the corresponding implementation lane is incomplete.</summary>
    NotImplemented
}
