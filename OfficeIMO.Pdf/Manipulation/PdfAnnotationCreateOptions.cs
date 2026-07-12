namespace OfficeIMO.Pdf;

/// <summary>Creates a standard annotation on an existing PDF page.</summary>
public sealed class PdfAnnotationCreateOptions {
    /// <summary>One-based target page number.</summary>
    public int PageNumber { get; set; } = 1;
    /// <summary>PDF annotation subtype such as Text, FreeText, Highlight, Line, Square, Circle, Polygon, PolyLine, Ink, Stamp, or Caret.</summary>
    public string Subtype { get; set; } = "Text";
    /// <summary>Annotation rectangle as left, bottom, right, top.</summary>
    public IReadOnlyList<double> Rectangle { get; set; } = new[] { 36D, 36D, 54D, 54D };
    /// <summary>Optional annotation contents.</summary>
    public string? Contents { get; set; }
    /// <summary>Optional author/title.</summary>
    public string? Title { get; set; }
    /// <summary>Optional stable /NM name.</summary>
    public string? Name { get; set; }
    /// <summary>Annotation flags. The default enables printing.</summary>
    public int Flags { get; set; } = 4;
    /// <summary>Optional RGB color.</summary>
    public IReadOnlyList<double>? Color { get; set; }
    /// <summary>Optional text-markup quadrilaterals.</summary>
    public IReadOnlyList<double>? QuadPoints { get; set; }
    /// <summary>Optional polygon/polyline vertices.</summary>
    public IReadOnlyList<double>? Vertices { get; set; }
    /// <summary>Optional line endpoints.</summary>
    public IReadOnlyList<double>? Line { get; set; }
    /// <summary>Optional ink paths.</summary>
    public IReadOnlyList<IReadOnlyList<double>>? InkPaths { get; set; }
    /// <summary>Optional line-start ending name.</summary>
    public string? LineStartEnding { get; set; }
    /// <summary>Optional line-end ending name.</summary>
    public string? LineEndEnding { get; set; }
    /// <summary>Optional icon or stamp name stored in /Name.</summary>
    public string? IconName { get; set; }
    /// <summary>Optional reply-parent annotation object number.</summary>
    public int? InReplyToObjectNumber { get; set; }
    /// <summary>Reply type, normally R or Group.</summary>
    public string? ReplyType { get; set; }
    /// <summary>Creates and links a popup annotation.</summary>
    public bool CreatePopup { get; set; }
    /// <summary>Popup rectangle. Defaults beside the parent rectangle.</summary>
    public IReadOnlyList<double>? PopupRectangle { get; set; }
    /// <summary>Initial popup open state.</summary>
    public bool PopupOpen { get; set; }
    /// <summary>Generates a normal appearance for supported visual subtypes.</summary>
    public bool GenerateAppearance { get; set; } = true;
    /// <summary>Preferred full-rewrite or append-only mutation mode.</summary>
    public PdfMutationExecutionPreference ExecutionPreference { get; set; } = PdfMutationExecutionPreference.Automatic;
}
