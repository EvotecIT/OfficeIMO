using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace OfficeIMO.Pdf;

/// <summary>
/// Options used when saving a Word document as PDF.
/// </summary>
public enum PdfPageOrientation {
    Portrait,
    Landscape
}

public class PdfSaveOptions {
    /// <summary>
    /// Gets or sets the page size. Defaults to <see cref="PageSizes.A4"/>.
    /// </summary>
    public PageSize PageSize { get; set; } = PageSizes.A4;

    /// <summary>
    /// Gets or sets the page orientation. Defaults to <see cref="PdfPageOrientation.Portrait"/>.
    /// </summary>
    public PdfPageOrientation Orientation { get; set; } = PdfPageOrientation.Portrait;
}

