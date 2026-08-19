namespace OfficeIMO.PowerPoint.Html;

/// <summary>Named PowerPoint-to-HTML output contracts.</summary>
public enum PowerPointHtmlExportProfile {
    /// <summary>Slides, notes, and slide content as accessible semantic HTML.</summary>
    SemanticSlides,

    /// <summary>Slides as positioned visual-review HTML backed by shared drawing primitives.</summary>
    VisualReview
}
