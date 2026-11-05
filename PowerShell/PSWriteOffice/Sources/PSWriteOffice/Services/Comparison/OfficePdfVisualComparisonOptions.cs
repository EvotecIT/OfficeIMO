using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Comparison;

/// <summary>Native comparison settings and distinct bounded input policies for the two PDF roles.</summary>
public sealed class OfficePdfVisualComparisonOptions {
    /// <summary>Native render/tolerance settings.</summary>
    public PdfVisualComparisonOptions Comparison { get; set; } = new();
    /// <summary>Policy used to load the expected PDF.</summary>
    public PdfLoadOptions? ReferenceLoadOptions { get; set; }
    /// <summary>Policy used to load the actual PDF.</summary>
    public PdfLoadOptions? DifferenceLoadOptions { get; set; }
}
