using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Builds foreground page content at absolute top-left page coordinates in the order items are added.
/// </summary>
public sealed class PdfPageCanvas {
    private readonly List<PdfCanvasItem> _items = new();

    internal IReadOnlyList<PdfCanvasItem> Items => _items;

    /// <summary>Adds text inside a fixed page rectangle using top-left page coordinates.</summary>
    public PdfPageCanvas Text(string text, double x, double y, double width, double height, double? fontSize = null, PdfColor? color = null, PdfAlign align = PdfAlign.Left, PdfStandardFont? font = null) {
        Guard.NotNull(text, nameof(text));
        return Text(new[] { TextRun.Normal(text, color, fontSize, font: font) }, x, y, width, height, color, align, fontSize);
    }

    /// <summary>Adds rich text runs inside a fixed page rectangle using top-left page coordinates.</summary>
    public PdfPageCanvas Text(IEnumerable<TextRun> runs, double x, double y, double width, double height, PdfColor? defaultColor = null, PdfAlign align = PdfAlign.Left, double? fontSize = null, double? lineHeight = null) {
        Guard.NotNull(runs, nameof(runs));
        Guard.NonNegative(x, nameof(x));
        Guard.NonNegative(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.ParagraphAlign(align, nameof(align), "Canvas text");
        if (fontSize.HasValue) {
            Guard.Positive(fontSize.Value, nameof(fontSize));
        }

        if (lineHeight.HasValue) {
            Guard.Positive(lineHeight.Value, nameof(lineHeight));
        }

        var snapshot = runs.ToList();
        if (snapshot.Count == 0) {
            throw new ArgumentException("Canvas text requires at least one text run.", nameof(runs));
        }

        _items.Add(new PdfCanvasTextItem(snapshot, x, y, width, height, defaultColor, align, fontSize, lineHeight));
        return this;
    }

    /// <summary>Adds a styled fixed-position text box using top-left page coordinates.</summary>
    public PdfPageCanvas TextBox(string text, double x, double y, double width, double height, PdfCanvasTextBoxStyle? style = null, double rotationAngle = 0D) {
        Guard.NotNull(text, nameof(text));
        return TextBox(new[] { TextRun.Normal(text) }, x, y, width, height, style, rotationAngle);
    }

    /// <summary>Adds styled rich text inside a fixed-position text box using top-left page coordinates.</summary>
    public PdfPageCanvas TextBox(IEnumerable<TextRun> runs, double x, double y, double width, double height, PdfCanvasTextBoxStyle? style = null, double rotationAngle = 0D) {
        Guard.NotNull(runs, nameof(runs));
        Guard.NonNegative(x, nameof(x));
        Guard.NonNegative(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        ValidateFiniteRotation(rotationAngle, nameof(rotationAngle), "Canvas text box rotation angle must be finite.");

        PdfCanvasTextBoxStyle textBoxStyle = style?.Clone() ?? new PdfCanvasTextBoxStyle();
        ValidateTextBoxInnerArea(width, height, textBoxStyle, nameof(style));

        var snapshot = runs.ToList();
        if (snapshot.Count == 0) {
            throw new ArgumentException("Canvas text box requires at least one text run.", nameof(runs));
        }

        _items.Add(new PdfCanvasTextBoxItem(ApplyTextBoxDefaults(snapshot, textBoxStyle), x, y, width, height, textBoxStyle, rotationAngle));
        return this;
    }

    /// <summary>Adds a shared drawing shape at fixed top-left page coordinates.</summary>
    public PdfPageCanvas Shape(OfficeShape shape, double x, double y, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null, double rotationAngle = 0D) {
        Guard.NotNull(shape, nameof(shape));
        Guard.NonNegative(x, nameof(x));
        Guard.NonNegative(y, nameof(y));
        ValidateFiniteRotation(rotationAngle, nameof(rotationAngle), "Canvas shape rotation angle must be finite.");
        ShapeBlock block = PdfDocument.CreateShapeBlock(CreateRotatedShape(shape, rotationAngle), PdfAlign.Left, spacingBefore: 0D, spacingAfter: 0D, style, linkUri, linkContents);
        _items.Add(new PdfCanvasShapeItem(block, x, y, rotationAngle));
        return this;
    }

    /// <summary>Adds a shared vector drawing inside a fixed page rectangle using top-left page coordinates.</summary>
    public PdfPageCanvas Drawing(OfficeDrawing drawing, double x, double y, double width, double height, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null, double rotationAngle = 0D) {
        Guard.NotNull(drawing, nameof(drawing));
        Guard.NonNegative(x, nameof(x));
        Guard.NonNegative(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        ValidateFiniteRotation(rotationAngle, nameof(rotationAngle), "Canvas drawing rotation angle must be finite.");

        DrawingBlock block = PdfDocument.CreateDrawingBlock(drawing, PdfAlign.Left, spacingBefore: 0D, spacingAfter: 0D, style, linkUri, linkContents);
        _items.Add(new PdfCanvasDrawingItem(block, x, y, width, height, rotationAngle));
        return this;
    }

    /// <summary>Adds a supported image at fixed top-left page coordinates.</summary>
    public PdfPageCanvas Image(byte[] imageBytes, double x, double y, double width, double height, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null, string? alternativeText = null, double rotationAngle = 0D, bool horizontalFlip = false, bool verticalFlip = false) {
        Guard.NotNullOrEmpty(imageBytes, nameof(imageBytes));
        Guard.NonNegative(x, nameof(x));
        Guard.NonNegative(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
        ValidateFiniteRotation(rotationAngle, nameof(rotationAngle), "Canvas image rotation angle must be finite.");

        PdfImageStyle? imageStyle = PdfDocument.CreateImageStyle(PdfAlign.Left, style?.ClipPath, style?.Fit, spacingBefore: 0D, spacingAfter: 0D, style, alternativeText);
        if (imageStyle != null) {
            PdfDocument.ValidateImageStyleForBox(imageStyle, width, height, nameof(style));
        }

        OfficeImageInfo imageInfo = PdfDocument.ValidateImageBytes(imageBytes);
        if (imageStyle != null) {
            PdfDocument.ValidateImageFitDimensions(imageInfo, imageStyle.Fit, nameof(style));
        }

        _items.Add(new PdfCanvasImageItem(new ImageBlock(imageBytes, width, height, imageInfo, imageStyle, linkUri, linkContents), x, y, rotationAngle, horizontalFlip, verticalFlip));
        return this;
    }

    /// <summary>Adds a fixed-position table inside a page rectangle using top-left page coordinates.</summary>
    public PdfPageCanvas Table(IEnumerable<string[]> rows, double x, double y, double width, double height, PdfTableStyle? style = null) {
        Guard.NotNull(rows, nameof(rows));
        Guard.NonNegative(x, nameof(x));
        Guard.NonNegative(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        _items.Add(new PdfCanvasTableItem(new TableBlock(rows, PdfAlign.Left, style), x, y, width, height));
        return this;
    }

    /// <summary>Adds a fixed-position rich table inside a page rectangle using top-left page coordinates.</summary>
    public PdfPageCanvas Table(IEnumerable<PdfTableCell[]> rows, double x, double y, double width, double height, PdfTableStyle? style = null) {
        Guard.NotNull(rows, nameof(rows));
        Guard.NonNegative(x, nameof(x));
        Guard.NonNegative(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        _items.Add(new PdfCanvasTableItem(new TableBlock(rows, PdfAlign.Left, style), x, y, width, height));
        return this;
    }

    private static OfficeShape CreateRotatedShape(OfficeShape shape, double rotationAngle) {
        if (rotationAngle == 0D) {
            return shape;
        }

        OfficeShape rotated = shape.Clone();
        OfficeTransform rotation = OfficeTransform.RotateDegrees(rotationAngle, rotated.Width / 2D, rotated.Height / 2D);
        rotated.Transform = rotated.Transform.HasValue ? rotated.Transform.Value.Then(rotation) : rotation;
        return rotated;
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<TextRun> ApplyTextBoxDefaults(List<TextRun> runs, PdfCanvasTextBoxStyle style) {
        var styled = new List<TextRun>(runs.Count);
        for (int i = 0; i < runs.Count; i++) {
            TextRun run = runs[i];
            bool applyTextColor = !run.Color.HasValue && style.TextColor.HasValue;
            bool applyFontSize = !run.FontSize.HasValue && style.FontSize.HasValue;
            bool applyFont = !run.Font.HasValue && style.Font.HasValue;
            if (!applyTextColor && !applyFontSize && !applyFont) {
                styled.Add(run);
                continue;
            }

            styled.Add(new TextRun(
                run.Text,
                run.Bold,
                run.Underline,
                applyTextColor ? style.TextColor : run.Color,
                run.Italic,
                run.Strike,
                applyFontSize ? style.FontSize : run.FontSize,
                applyFont ? style.Font : run.Font,
                run.LinkUri,
                run.LinkContents,
                run.Baseline,
                run.LinkDestinationName,
                run.TabLeader,
                run.TabAlignment,
                run.BackgroundColor));
        }

        return styled.AsReadOnly();
    }

    private static void ValidateTextBoxInnerArea(double width, double height, PdfCanvasTextBoxStyle style, string paramName) {
        if (style.PaddingX * 2D >= width || style.PaddingY * 2D >= height) {
            throw new ArgumentException("Canvas text box padding must leave a positive text area.", paramName);
        }

        if (style.CornerRadius > Math.Min(width, height) / 2D) {
            throw new ArgumentOutOfRangeException(paramName, "Canvas text box corner radius cannot exceed half of the text box width or height.");
        }
    }

    private static void ValidateFiniteRotation(double value, string paramName, string message) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, message);
        }
    }
}

internal sealed class PdfCanvasBlock : IPdfBlock {
    public PdfCanvasBlock(IEnumerable<PdfCanvasItem> items) {
        Guard.NotNull(items, nameof(items));
        Items = items.ToList().AsReadOnly();
    }

    public IReadOnlyList<PdfCanvasItem> Items { get; }
}

internal abstract class PdfCanvasItem {
    protected PdfCanvasItem(double x, double y) {
        X = x;
        Y = y;
    }

    public double X { get; }
    public double Y { get; }
}

internal sealed class PdfCanvasTextItem : PdfCanvasItem {
    public PdfCanvasTextItem(IReadOnlyList<TextRun> runs, double x, double y, double width, double height, PdfColor? defaultColor, PdfAlign align, double? fontSize, double? lineHeight)
        : base(x, y) {
        Runs = runs;
        Width = width;
        Height = height;
        DefaultColor = defaultColor;
        Align = align;
        FontSize = fontSize;
        LineHeight = lineHeight;
    }

    public IReadOnlyList<TextRun> Runs { get; }
    public double Width { get; }
    public double Height { get; }
    public PdfColor? DefaultColor { get; }
    public PdfAlign Align { get; }
    public double? FontSize { get; }
    public double? LineHeight { get; }
}

internal sealed class PdfCanvasTextBoxItem : PdfCanvasItem {
    public PdfCanvasTextBoxItem(IReadOnlyList<TextRun> runs, double x, double y, double width, double height, PdfCanvasTextBoxStyle style, double rotationAngle)
        : base(x, y) {
        Runs = runs;
        Width = width;
        Height = height;
        Style = style.Clone();
        RotationAngle = rotationAngle;
    }

    public IReadOnlyList<TextRun> Runs { get; }
    public double Width { get; }
    public double Height { get; }
    public PdfCanvasTextBoxStyle Style { get; }
    public double RotationAngle { get; }
}

internal sealed class PdfCanvasShapeItem : PdfCanvasItem {
    public PdfCanvasShapeItem(ShapeBlock block, double x, double y, double rotationAngle)
        : base(x, y) {
        Block = block;
        RotationAngle = rotationAngle;
    }

    public ShapeBlock Block { get; }
    public double RotationAngle { get; }
}

internal sealed class PdfCanvasDrawingItem : PdfCanvasItem {
    public PdfCanvasDrawingItem(DrawingBlock block, double x, double y, double width, double height, double rotationAngle)
        : base(x, y) {
        Block = block;
        Width = width;
        Height = height;
        RotationAngle = rotationAngle;
    }

    public DrawingBlock Block { get; }
    public double Width { get; }
    public double Height { get; }
    public double RotationAngle { get; }
}

internal sealed class PdfCanvasImageItem : PdfCanvasItem {
    public PdfCanvasImageItem(ImageBlock block, double x, double y, double rotationAngle, bool horizontalFlip, bool verticalFlip)
        : base(x, y) {
        Block = block;
        RotationAngle = rotationAngle;
        HorizontalFlip = horizontalFlip;
        VerticalFlip = verticalFlip;
    }

    public ImageBlock Block { get; }
    public double RotationAngle { get; }
    public bool HorizontalFlip { get; }
    public bool VerticalFlip { get; }
}

internal sealed class PdfCanvasTableItem : PdfCanvasItem {
    public PdfCanvasTableItem(TableBlock block, double x, double y, double width, double height)
        : base(x, y) {
        Block = block;
        Width = width;
        Height = height;
    }

    public TableBlock Block { get; }
    public double Width { get; }
    public double Height { get; }
}
