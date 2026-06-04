using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Sets the document-wide default style for panel paragraphs.</summary>
    public PdfDocument DefaultPanelStyle(PanelStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultPanelStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default style for horizontal rules.</summary>
    public PdfDocument DefaultHorizontalRuleStyle(PdfHorizontalRuleStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultHorizontalRuleStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default image placement style.</summary>
    public PdfDocument DefaultImageStyle(PdfImageStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultImageStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default placement style for OfficeIMO.Drawing-backed flow objects.</summary>
    public PdfDocument DefaultDrawingStyle(PdfDrawingStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultDrawingStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default row/column layout style.</summary>
    public PdfDocument DefaultRowStyle(PdfRowStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultRowStyle = style;
        return this;
    }
}
