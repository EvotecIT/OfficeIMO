namespace OfficeIMO.ChartForgeX;

/// <summary>Identifies the native Visio projection strategy selected for a CFX artifact.</summary>
public enum OfficeVisioVisualProjectionKind {
    /// <summary>A native editable graph diagram.</summary>
    Graph,
    /// <summary>A native editable graph diagram specialized for flow semantics.</summary>
    FlowGraph,
    /// <summary>A native editable sequence diagram.</summary>
    Sequence
}

/// <summary>Identifies the severity of a native Visio projection diagnostic.</summary>
public enum OfficeVisioVisualDiagnosticSeverity {
    /// <summary>Informational fidelity detail.</summary>
    Information,
    /// <summary>A source semantic could not be represented exactly.</summary>
    Warning
}

/// <summary>Identifies a stable native Visio projection diagnostic category.</summary>
public enum OfficeVisioVisualDiagnosticCode {
    /// <summary>A renderer-only watermark was not projected.</summary>
    WatermarkNotProjected,
    /// <summary>Native layout recomputed prepared coordinates.</summary>
    LayoutRecomputed,
    /// <summary>A layout token was normalized.</summary>
    LayoutNormalized,
    /// <summary>A direction token was normalized.</summary>
    DirectionNormalized,
    /// <summary>A line style was normalized.</summary>
    LineStyleNormalized,
    /// <summary>A color could not be represented natively.</summary>
    ColorNotProjected,
    /// <summary>A tooltip was retained as Shape Data.</summary>
    TooltipRetainedAsShapeData,
    /// <summary>A tooltip could not be projected.</summary>
    TooltipNotProjected,
    /// <summary>An extension key was renamed for Visio Shape Data.</summary>
    ExtensionKeyRenamed,
    /// <summary>A metric name was renamed for Visio Shape Data.</summary>
    MetricNameRenamed,
    /// <summary>A detail field was renamed for Visio Shape Data.</summary>
    DetailFieldRenamed,
    /// <summary>A source id was remapped to avoid a native collision.</summary>
    IdRemapped,
    /// <summary>A graph group was not projected.</summary>
    GroupNotProjected,
    /// <summary>An annotation was not projected.</summary>
    AnnotationNotProjected,
    /// <summary>Artifact-level extensions were not projected.</summary>
    ExtensionsNotProjected,
    /// <summary>Accessibility semantics were not projected.</summary>
    AccessibilityNotProjected,
    /// <summary>Artifact presentation semantics were not projected as native page constructs.</summary>
    PresentationNotProjected,
    /// <summary>A reusable scenario was retained only in the semantic envelope.</summary>
    ScenarioNotProjected,
    /// <summary>Portable artwork was retained only in the semantic envelope.</summary>
    ArtworkNotProjected,
    /// <summary>Prepared edge geometry or advanced edge presentation was recomputed or omitted.</summary>
    EdgePresentationNormalized,
    /// <summary>Port attachment semantics were normalized.</summary>
    PortAttachmentNormalized,
    /// <summary>Endpoint labels were not rendered.</summary>
    EndpointLabelsNotRendered,
    /// <summary>A sequence note placement or span was normalized.</summary>
    NoteNormalized,
    /// <summary>Shape Data was disabled for a semantic that could otherwise be retained there.</summary>
    ShapeDataDisabled,
    /// <summary>A source hyperlink was not projected as an active native hyperlink.</summary>
    HyperlinkNotProjected,
    /// <summary>A semantic has no more specific diagnostic category.</summary>
    SemanticLoss
}

/// <summary>Identifies the source entity associated with a projection diagnostic.</summary>
public enum OfficeVisioVisualEntityKind {
    /// <summary>The complete artifact.</summary>
    Artifact,
    /// <summary>A group or container.</summary>
    Group,
    /// <summary>A graph node.</summary>
    Node,
    /// <summary>A graph edge.</summary>
    Edge,
    /// <summary>A sequence participant.</summary>
    Participant,
    /// <summary>A sequence message.</summary>
    Message,
    /// <summary>An annotation.</summary>
    Annotation,
    /// <summary>A node detail row.</summary>
    Detail,
    /// <summary>A node port.</summary>
    Port
}

/// <summary>Describes one machine-readable native Visio projection fidelity decision.</summary>
public sealed class OfficeVisioVisualDiagnostic {
    internal OfficeVisioVisualDiagnostic(
        OfficeVisioVisualDiagnosticCode code,
        OfficeVisioVisualDiagnosticSeverity severity,
        OfficeVisioVisualEntityKind entityKind,
        string? entityId,
        string? feature,
        string message) {
        Code = code;
        Severity = severity;
        EntityKind = entityKind;
        EntityId = entityId;
        Feature = feature;
        Message = message;
    }

    /// <summary>Gets the stable diagnostic code.</summary>
    public OfficeVisioVisualDiagnosticCode Code { get; }
    /// <summary>Gets the diagnostic severity.</summary>
    public OfficeVisioVisualDiagnosticSeverity Severity { get; }
    /// <summary>Gets the associated entity kind.</summary>
    public OfficeVisioVisualEntityKind EntityKind { get; }
    /// <summary>Gets the associated source entity id.</summary>
    public string? EntityId { get; }
    /// <summary>Gets the semantic feature name.</summary>
    public string? Feature { get; }
    /// <summary>Gets the human-readable explanation.</summary>
    public string Message { get; }
}
