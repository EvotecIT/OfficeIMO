using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Sets the document-wide default style for panel paragraphs.</summary>
    internal PdfDocument DefaultPanelStyle(PanelStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultPanelStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default style for horizontal rules.</summary>
    internal PdfDocument DefaultHorizontalRuleStyle(PdfHorizontalRuleStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultHorizontalRuleStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default image placement style.</summary>
    internal PdfDocument DefaultImageStyle(PdfImageStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultImageStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default placement style for OfficeIMO.Drawing-backed flow objects.</summary>
    internal PdfDocument DefaultDrawingStyle(PdfDrawingStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultDrawingStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default row/column layout style.</summary>
    internal PdfDocument DefaultRowStyle(PdfRowStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultRowStyle = style;
        return this;
    }
}
