using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDoc {
    /// <summary>Sets the document-wide default style for panel paragraphs.</summary>
    public PdfDoc DefaultPanelStyle(PanelStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultPanelStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default style for horizontal rules.</summary>
    public PdfDoc DefaultHorizontalRuleStyle(PdfHorizontalRuleStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultHorizontalRuleStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default image placement style.</summary>
    public PdfDoc DefaultImageStyle(PdfImageStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultImageStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default placement style for OfficeIMO.Drawing-backed flow objects.</summary>
    public PdfDoc DefaultDrawingStyle(PdfDrawingStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultDrawingStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default row/column layout style.</summary>
    public PdfDoc DefaultRowStyle(PdfRowStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultRowStyle = style;
        return this;
    }
}
