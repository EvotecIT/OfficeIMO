namespace OfficeIMO.Pdf;

/// <summary>Destructively cropped PDF bytes plus planner and raster readback evidence.</summary>
public sealed class PdfDestructiveCropResult {
    private readonly byte[] _pdf;
    internal PdfDestructiveCropResult(byte[] pdf, PdfMutationPlan pageTreePlan, PdfMutationPlan contentPlan, PdfRewritePreservationReport preservation, IReadOnlyList<PdfPageRenderResult> renders) { _pdf = (byte[])pdf.Clone(); PageTreePlan = pageTreePlan; ContentPlan = contentPlan; PreservationReport = preservation; Renders = renders; }
    /// <summary>Plan authorizing page-boundary and coordinate changes.</summary>
    public PdfMutationPlan PageTreePlan { get; }
    /// <summary>Plan authorizing destructive content replacement.</summary>
    public PdfMutationPlan ContentPlan { get; }
    /// <summary>Non-target structure preservation proof.</summary>
    public PdfRewritePreservationReport PreservationReport { get; }
    /// <summary>Raster artifacts used to replace selected pages.</summary>
    public IReadOnlyList<PdfPageRenderResult> Renders { get; }
    /// <summary>Returns a defensive copy of the destructively cropped PDF.</summary>
    public byte[] ToBytes() => (byte[])_pdf.Clone();
    /// <summary>Opens the destructively cropped artifact.</summary>
    public PdfDocument ToDocument() => PdfDocument.Open(_pdf);
}
