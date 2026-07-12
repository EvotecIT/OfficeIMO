using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

/// <summary>
/// Summary of a semantic PowerPoint HTML import.
/// </summary>
public sealed class HtmlToPowerPointResult : HtmlConversionResult<PptCore.PowerPointPresentation> {
    internal HtmlToPowerPointResult(PptCore.PowerPointPresentation presentation) : base(presentation) { }

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
    /// Number of merged table ranges restored from HTML row and column spans.
    /// </summary>
    public int MergedRanges { get; internal set; }

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

}
