using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// Markdown-aware figure styling for images and generated visuals in PDF export.
/// </summary>
public sealed class MarkdownPdfFigureStyle {
    private PdfCore.PanelStyle? _panelStyle;
    private PdfCore.PdfImageStyle? _imageStyle;
    private PdfCore.PdfDrawingStyle? _drawingStyle;
    private PdfCore.PdfColor? _captionColor;
    private PdfCore.PdfColor? _placeholderColor;
    private double? _captionFontSize;
    private PdfCore.PdfAlign _captionAlign = PdfCore.PdfAlign.Center;

    /// <summary>Optional panel chrome used for unresolved visual placeholders. Media figures render directly in document flow.</summary>
    public PdfCore.PanelStyle? PanelStyle {
        get => _panelStyle?.Clone();
        set => _panelStyle = value?.Clone();
    }

    /// <summary>Image placement defaults used for Markdown image blocks and image-only paragraphs.</summary>
    public PdfCore.PdfImageStyle? ImageStyle {
        get => _imageStyle?.Clone();
        set => _imageStyle = value?.Clone();
    }

    /// <summary>Drawing placement defaults used for generated visual blocks such as Chart.js-compatible chart fences.</summary>
    public PdfCore.PdfDrawingStyle? DrawingStyle {
        get => _drawingStyle?.Clone();
        set => _drawingStyle = value?.Clone();
    }

    /// <summary>Caption text color.</summary>
    public PdfCore.PdfColor? CaptionColor {
        get => _captionColor;
        set => _captionColor = value;
    }

    /// <summary>Caption font size.</summary>
    public double? CaptionFontSize {
        get => _captionFontSize;
        set => _captionFontSize = ValidateOptionalPositive(value, nameof(CaptionFontSize));
    }

    /// <summary>Caption paragraph alignment.</summary>
    public PdfCore.PdfAlign CaptionAlign {
        get => _captionAlign;
        set {
            if (value != PdfCore.PdfAlign.Left && value != PdfCore.PdfAlign.Center && value != PdfCore.PdfAlign.Right) {
                throw new ArgumentOutOfRangeException(nameof(value), "Markdown PDF figure captions support Left, Center, or Right alignment.");
            }

            _captionAlign = value;
        }
    }

    /// <summary>Text color used by unresolved image and visual placeholders.</summary>
    public PdfCore.PdfColor? PlaceholderColor {
        get => _placeholderColor;
        set => _placeholderColor = value;
    }

    /// <summary>Creates a copy of this figure style.</summary>
    public MarkdownPdfFigureStyle Clone() => new MarkdownPdfFigureStyle {
        PanelStyle = _panelStyle,
        ImageStyle = _imageStyle,
        DrawingStyle = _drawingStyle,
        CaptionColor = _captionColor,
        CaptionFontSize = _captionFontSize,
        CaptionAlign = _captionAlign,
        PlaceholderColor = _placeholderColor
    };

    internal PdfCore.PanelStyle? PanelStyleSnapshot => _panelStyle?.Clone();
    internal PdfCore.PdfImageStyle ImageStyleSnapshot => (_imageStyle ?? CreateDefaultImageStyle()).Clone();
    internal PdfCore.PdfDrawingStyle DrawingStyleSnapshot => (_drawingStyle ?? CreateDefaultDrawingStyle()).Clone();
    internal PdfCore.PdfColor CaptionColorSnapshot => _captionColor ?? PdfCore.PdfColor.FromRgb(71, 85, 105);
    internal double CaptionFontSizeSnapshot => _captionFontSize ?? 9D;
    internal PdfCore.PdfAlign CaptionAlignSnapshot => _captionAlign;
    internal PdfCore.PdfColor PlaceholderColorSnapshot => _placeholderColor ?? PdfCore.PdfColor.FromRgb(100, 116, 139);

    internal static MarkdownPdfFigureStyle Plain() => new MarkdownPdfFigureStyle {
        ImageStyle = CreateDefaultImageStyle(),
        DrawingStyle = CreateDefaultDrawingStyle(),
        CaptionColor = PdfCore.PdfColor.FromRgb(71, 85, 105),
        PlaceholderColor = PdfCore.PdfColor.FromRgb(100, 116, 139),
        CaptionFontSize = 9D,
        CaptionAlign = PdfCore.PdfAlign.Center
    };

    internal static MarkdownPdfFigureStyle Framed(
        PdfCore.PdfColor background,
        PdfCore.PdfColor border,
        PdfCore.PdfColor caption,
        PdfCore.PdfColor placeholder,
        double borderWidth = 0.5D) => new MarkdownPdfFigureStyle {
            PanelStyle = new PdfCore.PanelStyle {
                Background = background,
                BorderColor = border,
                BorderWidth = borderWidth,
                PaddingX = 10,
                PaddingY = 8,
                SpacingBefore = 5,
                SpacingAfter = 10,
                KeepTogether = true
            },
            ImageStyle = CreateDefaultImageStyle(),
            DrawingStyle = CreateDefaultDrawingStyle(),
            CaptionColor = caption,
            PlaceholderColor = placeholder,
            CaptionFontSize = 9D,
            CaptionAlign = PdfCore.PdfAlign.Center
        };

    private static PdfCore.PdfImageStyle CreateDefaultImageStyle() => new PdfCore.PdfImageStyle {
        Align = PdfCore.PdfAlign.Center,
        SpacingBefore = 2,
        SpacingAfter = 4,
        ScaleDownToFit = true,
        KeepWithNext = true
    };

    private static PdfCore.PdfDrawingStyle CreateDefaultDrawingStyle() => new PdfCore.PdfDrawingStyle {
        Align = PdfCore.PdfAlign.Center,
        SpacingBefore = 2,
        SpacingAfter = 4,
        KeepWithNext = true
    };

    private static double? ValidateOptionalPositive(double? value, string propertyName) {
        if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value) || value.Value <= 0)) {
            throw new ArgumentOutOfRangeException(propertyName, "Markdown PDF figure style sizes must be positive finite values.");
        }

        return value;
    }
}
