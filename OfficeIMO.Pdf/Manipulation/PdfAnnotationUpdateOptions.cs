namespace OfficeIMO.Pdf;

/// <summary>Updates for a single indirect PDF annotation during a safe full rewrite.</summary>
public sealed class PdfAnnotationUpdateOptions {
    /// <summary>Preferred mutation mode. Automatic uses a full rewrite when safe and append-only when required.</summary>
    public PdfMutationExecutionPreference ExecutionPreference { get; set; } = PdfMutationExecutionPreference.Automatic;

    /// <summary>Replacement /Contents text. Null leaves the value unchanged.</summary>
    public string? Contents { get; set; }

    /// <summary>Replacement /T title text. Null leaves the value unchanged.</summary>
    public string? Title { get; set; }

    /// <summary>Replacement /NM annotation name. Null leaves the value unchanged.</summary>
    public string? Name { get; set; }

    /// <summary>Replacement /F annotation flags. Null leaves the value unchanged.</summary>
    public int? Flags { get; set; }

    /// <summary>Replacement RGB /C color values in the 0..1 range. Null leaves the value unchanged.</summary>
    public IReadOnlyList<double>? Color { get; set; }

    /// <summary>Remove /A and /AA action dictionaries from the annotation.</summary>
    public bool RemoveActions { get; set; }

    /// <summary>Replacement annotation rectangle as left, bottom, right, top coordinates.</summary>
    public IReadOnlyList<double>? Rectangle { get; set; }

    /// <summary>Replacement text-markup quadrilaterals as groups of eight coordinates.</summary>
    public IReadOnlyList<double>? QuadPoints { get; set; }

    /// <summary>Replacement polygon or polyline vertices as x/y coordinate pairs.</summary>
    public IReadOnlyList<double>? Vertices { get; set; }

    /// <summary>Replacement line endpoints as x1, y1, x2, y2.</summary>
    public IReadOnlyList<double>? Line { get; set; }

    /// <summary>Replacement ink paths; every path is an x/y coordinate sequence.</summary>
    public IReadOnlyList<IReadOnlyList<double>>? InkPaths { get; set; }

    /// <summary>Replacement line-start ending name, for example None, OpenArrow, or ClosedArrow.</summary>
    public string? LineStartEnding { get; set; }

    /// <summary>Replacement line-end ending name.</summary>
    public string? LineEndEnding { get; set; }

    /// <summary>Replacement indirect annotation object used as /IRT reply parent.</summary>
    public int? InReplyToObjectNumber { get; set; }

    /// <summary>Replacement reply type stored in /RT, normally R or Group.</summary>
    public string? ReplyType { get; set; }

    /// <summary>Replacement popup open state for a popup annotation or linked /Popup dictionary.</summary>
    public bool? PopupOpen { get; set; }

    /// <summary>Replacement linked popup rectangle as left, bottom, right, top coordinates.</summary>
    public IReadOnlyList<double>? PopupRectangle { get; set; }

    /// <summary>Regenerates a supported normal appearance stream after applying the update.</summary>
    public bool RegenerateAppearance { get; set; }

    /// <summary>
    /// Allows append-only updates even though prior revisions retain replaced annotation data.
    /// Disabled by default because append-only output is not a sanitization boundary.
    /// </summary>
    public bool AllowResidualDataInAppendOnly { get; set; }
}
