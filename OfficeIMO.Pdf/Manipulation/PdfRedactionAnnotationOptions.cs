using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Options for a standard PDF /Redact review annotation.</summary>
public sealed class PdfRedactionAnnotationOptions {
    /// <summary>Creates options for a required review region.</summary>
    public PdfRedactionAnnotationOptions(PdfRedactionRegion region) {
        Region = region ?? throw new ArgumentNullException(nameof(region));
    }

    /// <summary>Region recorded by the annotation.</summary>
    public PdfRedactionRegion Region { get; }
    /// <summary>Optional review note stored in /Contents.</summary>
    public string? Contents { get; set; }
    /// <summary>Optional author stored in /T.</summary>
    public string? Author { get; set; }
    /// <summary>Optional stable annotation name stored in /NM.</summary>
    public string? Name { get; set; }
    /// <summary>Annotation color. Defaults to red.</summary>
    public IReadOnlyList<double> Color { get; set; } = new[] { 1D, 0D, 0D };
    /// <summary>Maximum canonical rectangles authored from the region, from 1 through 64. Defaults to 16 while annotation creation performs one bounded rewrite per rectangle.</summary>
    public int MaximumAnnotations { get; set; } = 16;
    /// <summary>Cooperatively cancels between annotation rewrites.</summary>
    public CancellationToken CancellationToken { get; set; }
}
