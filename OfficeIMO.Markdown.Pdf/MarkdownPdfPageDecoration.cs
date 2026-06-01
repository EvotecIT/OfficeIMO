using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// Page-level visual treatment applied by Markdown PDF themes through the shared OfficeIMO.Pdf engine.
/// </summary>
public sealed class MarkdownPdfPageDecoration {
    private PdfCore.PdfPageBorder? _pageBorder;
    private List<PdfCore.PdfPageBackgroundShape>? _backgroundShapes;
    private MarkdownPdfPageDecorationKind _kind;

    /// <summary>Optional page background color applied before page content is rendered.</summary>
    public PdfCore.PdfColor? BackgroundColor { get; set; }

    /// <summary>Optional page border applied before page content is rendered.</summary>
    public PdfCore.PdfPageBorder? PageBorder {
        get => _pageBorder?.Clone();
        set => _pageBorder = value?.Clone();
    }

    /// <summary>Optional page background shapes applied before page content is rendered.</summary>
    public IReadOnlyList<PdfCore.PdfPageBackgroundShape>? BackgroundShapes {
        get => CloneBackgroundShapes(_backgroundShapes);
        set => _backgroundShapes = CloneBackgroundShapes(value);
    }

    /// <summary>
    /// When true, explicit low-level PdfOptions values take precedence over this theme decoration.
    /// </summary>
    public bool RespectExistingPdfOptions { get; set; } = true;

    /// <summary>Creates the built-in report page decoration profile.</summary>
    public static MarkdownPdfPageDecoration Report() => new MarkdownPdfPageDecoration {
        _kind = MarkdownPdfPageDecorationKind.Report,
        BackgroundColor = PdfCore.PdfColor.FromRgb(248, 250, 252),
        PageBorder = new PdfCore.PdfPageBorder {
            Color = PdfCore.PdfColor.FromRgb(147, 197, 253),
            Width = 0.55,
            Inset = 34,
            Opacity = 0.35
        }
    };

    /// <summary>Creates the built-in technical document page decoration profile.</summary>
    public static MarkdownPdfPageDecoration TechnicalDocument() => new MarkdownPdfPageDecoration {
        _kind = MarkdownPdfPageDecorationKind.TechnicalDocument,
        BackgroundColor = PdfCore.PdfColor.White,
        PageBorder = new PdfCore.PdfPageBorder {
            Color = PdfCore.PdfColor.FromRgb(203, 213, 225),
            Width = 0.5,
            Inset = 36,
            Opacity = 0.42
        }
    };

    /// <summary>Adds a reusable page background shape.</summary>
    public MarkdownPdfPageDecoration AddBackgroundShape(PdfCore.PdfPageBackgroundShape shape) {
        if (shape == null) {
            throw new ArgumentNullException(nameof(shape));
        }

        _backgroundShapes ??= new List<PdfCore.PdfPageBackgroundShape>();
        _backgroundShapes.Add(shape.Clone());
        return this;
    }

    /// <summary>Removes all page background shapes from this decoration profile.</summary>
    public MarkdownPdfPageDecoration ClearBackgroundShapes() {
        _backgroundShapes?.Clear();
        return this;
    }

    /// <summary>Creates a detached copy.</summary>
    public MarkdownPdfPageDecoration Clone() => new MarkdownPdfPageDecoration {
        BackgroundColor = BackgroundColor,
        PageBorder = _pageBorder,
        BackgroundShapes = _backgroundShapes,
        RespectExistingPdfOptions = RespectExistingPdfOptions,
        _kind = _kind
    };

    internal void Apply(PdfCore.PdfDoc pdf, PdfCore.PdfOptions options) {
        if (pdf == null) {
            throw new ArgumentNullException(nameof(pdf));
        }

        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        if (BackgroundColor.HasValue && ShouldApplyBackgroundColor(options)) {
            pdf.Background(BackgroundColor.Value);
        }

        IReadOnlyList<PdfCore.PdfPageBackgroundShape> shapes = CreateEffectiveBackgroundShapes(options);
        if (shapes != null && shapes.Count > 0 && ShouldApplyBackgroundShapes(options)) {
            pdf.BackgroundShapes(shapes);
        }

        PdfCore.PdfPageBorder? border = PageBorder;
        if (border != null && ShouldApplyPageBorder(options)) {
            pdf.PageBorder(border);
        }
    }

    private bool ShouldApplyBackgroundColor(PdfCore.PdfOptions options) =>
        !RespectExistingPdfOptions || !options.BackgroundColor.HasValue;

    private bool ShouldApplyBackgroundShapes(PdfCore.PdfOptions options) =>
        !RespectExistingPdfOptions || options.PageBackgroundShapes == null || options.PageBackgroundShapes.Count == 0;

    private bool ShouldApplyPageBorder(PdfCore.PdfOptions options) =>
        !RespectExistingPdfOptions || options.PageBorder == null;

    private IReadOnlyList<PdfCore.PdfPageBackgroundShape> CreateEffectiveBackgroundShapes(PdfCore.PdfOptions options) {
        var shapes = new List<PdfCore.PdfPageBackgroundShape>();
        switch (_kind) {
            case MarkdownPdfPageDecorationKind.Report:
                shapes.AddRange(CreateReportBackgroundShapes(options));
                break;
            case MarkdownPdfPageDecorationKind.TechnicalDocument:
                shapes.AddRange(CreateTechnicalDocumentBackgroundShapes(options));
                break;
        }

        if (_backgroundShapes != null) {
            foreach (PdfCore.PdfPageBackgroundShape shape in _backgroundShapes) {
                shapes.Add(shape.Clone());
            }
        }

        return shapes;
    }

    private static IEnumerable<PdfCore.PdfPageBackgroundShape> CreateReportBackgroundShapes(PdfCore.PdfOptions options) {
        yield return PdfCore.PdfPageBackgroundShape.TopBand(
            options.PageWidth,
            options.PageHeight,
            110,
            fill: PdfCore.PdfColor.FromRgb(239, 246, 255),
            insetX: 34,
            offsetY: 42,
            cornerRadius: 20,
            stroke: PdfCore.PdfColor.FromRgb(191, 219, 254),
            strokeWidth: 0.55,
            fillOpacity: 0.58,
            strokeOpacity: 0.45,
            fillGradient: HorizontalGradient(PdfCore.PdfColor.FromRgb(239, 246, 255), PdfCore.PdfColor.White));

        yield return PdfCore.PdfPageBackgroundShape.RightBand(
            options.PageWidth,
            options.PageHeight,
            54,
            fill: PdfCore.PdfColor.FromRgb(219, 234, 254),
            insetY: 58,
            offsetX: 34,
            cornerRadius: 20,
            fillOpacity: 0.24);

        yield return PdfCore.PdfPageBackgroundShape.Ellipse(
            Math.Max(0, options.PageWidth - 160),
            Math.Max(0, options.PageHeight - 152),
            112,
            112,
            fill: PdfCore.PdfColor.FromRgb(191, 219, 254),
            fillOpacity: 0.22);
    }

    private static IEnumerable<PdfCore.PdfPageBackgroundShape> CreateTechnicalDocumentBackgroundShapes(PdfCore.PdfOptions options) {
        yield return PdfCore.PdfPageBackgroundShape.TopBand(
            options.PageWidth,
            options.PageHeight,
            78,
            fill: PdfCore.PdfColor.FromRgb(248, 250, 252),
            insetX: 36,
            offsetY: 36,
            cornerRadius: 14,
            stroke: PdfCore.PdfColor.FromRgb(226, 232, 240),
            strokeWidth: 0.45,
            fillOpacity: 0.82,
            strokeOpacity: 0.55,
            fillGradient: HorizontalGradient(PdfCore.PdfColor.FromRgb(248, 250, 252), PdfCore.PdfColor.White));

        yield return PdfCore.PdfPageBackgroundShape.LeftBand(
            options.PageWidth,
            options.PageHeight,
            8,
            fill: PdfCore.PdfColor.FromRgb(9, 105, 218),
            insetY: 54,
            offsetX: 36,
            cornerRadius: 4,
            fillOpacity: 0.18);

        yield return PdfCore.PdfPageBackgroundShape.Rectangle(
            options.PageWidth - 120,
            options.PageHeight - 88,
            48,
            6,
            fill: PdfCore.PdfColor.FromRgb(9, 105, 218),
            fillOpacity: 0.18);
    }

    private static OfficeLinearGradient HorizontalGradient(PdfCore.PdfColor startColor, PdfCore.PdfColor endColor) =>
        OfficeLinearGradient.Horizontal(startColor.ToOfficeColor(), endColor.ToOfficeColor());

    private static List<PdfCore.PdfPageBackgroundShape>? CloneBackgroundShapes(IEnumerable<PdfCore.PdfPageBackgroundShape>? shapes) {
        if (shapes == null) {
            return null;
        }

        var cloned = new List<PdfCore.PdfPageBackgroundShape>();
        foreach (PdfCore.PdfPageBackgroundShape shape in shapes) {
            if (shape == null) {
                throw new ArgumentNullException(nameof(shapes), "Markdown PDF page decoration background shapes cannot contain null entries.");
            }

            cloned.Add(shape.Clone());
        }

        return cloned;
    }

    private enum MarkdownPdfPageDecorationKind {
        None,
        Report,
        TechnicalDocument
    }
}
