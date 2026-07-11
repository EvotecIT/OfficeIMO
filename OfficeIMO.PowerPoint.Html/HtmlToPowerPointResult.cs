using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>
/// Summary of a semantic PowerPoint HTML import.
/// </summary>
public sealed class HtmlToPowerPointResult {
    internal HtmlToPowerPointResult(PptCore.PowerPointPresentation presentation) {
        Presentation = presentation ?? throw new ArgumentNullException(nameof(presentation));
    }

    /// <summary>
    /// Imported presentation.
    /// </summary>
    public PptCore.PowerPointPresentation Presentation { get; }

    /// <summary>
    /// Number of imported slides.
    /// </summary>
    public int Slides { get; internal set; }

    /// <summary>
    /// Number of imported text boxes.
    /// </summary>
    public int TextBoxes { get; internal set; }

    /// <summary>
    /// Number of imported tables.
    /// </summary>
    public int Tables { get; internal set; }

    /// <summary>
    /// Number of imported pictures.
    /// </summary>
    public int Pictures { get; internal set; }

    /// <summary>
    /// Number of chart inventory items restored as native charts with reconstructed placeholder data.
    /// </summary>
    public int Charts { get; internal set; }

    /// <summary>
    /// Number of presenter note blocks restored.
    /// </summary>
    public int Notes { get; internal set; }

    /// <summary>
    /// Import diagnostics for skipped or approximate rich content.
    /// </summary>
    public IList<string> Diagnostics { get; } = new List<string>();
}
