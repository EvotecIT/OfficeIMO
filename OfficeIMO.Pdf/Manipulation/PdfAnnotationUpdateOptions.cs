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
}
