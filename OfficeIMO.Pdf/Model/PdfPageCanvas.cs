using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Builds foreground page content at absolute top-left page coordinates in the order items are added.
/// </summary>
public sealed class PdfPageCanvas {
    private readonly List<PdfCanvasItem> _items = new();
    private readonly bool _allowOutOfPageCoordinates;

    /// <summary>Creates an empty absolute-positioning page canvas.</summary>
    public PdfPageCanvas() {
    }

    private PdfPageCanvas(bool allowOutOfPageCoordinates) {
        _allowOutOfPageCoordinates = allowOutOfPageCoordinates;
    }

    internal IReadOnlyList<PdfCanvasItem> Items => _items;

    /// <summary>Adds a document outline entry targeting an absolute top-left page coordinate.</summary>
    /// <param name="title">Visible outline title.</param>
    /// <param name="level">One-based outline hierarchy level.</param>
    /// <param name="y">Top coordinate in page points.</param>
    public PdfPageCanvas Outline(string title, int level, double y) {
        Guard.NotNull(title, nameof(title));
        if (string.IsNullOrWhiteSpace(title)) {
            throw new ArgumentException("Canvas outline titles cannot be empty or whitespace.", nameof(title));
        }

        if (level <= 0) {
            throw new ArgumentOutOfRangeException(nameof(level), "Canvas outline levels must be positive.");
        }

        ValidateCanvasCoordinate(y, nameof(y));
        _items.Add(new PdfCanvasOutlineItem(title.Trim(), level, y));
        return this;
    }

    /// <summary>Groups absolute canvas content as one tagged figure with alternative text.</summary>
    public PdfPageCanvas Figure(string alternativeText, Action<PdfPageCanvas> build) {
        Guard.NotNullOrWhiteSpace(alternativeText, nameof(alternativeText));
        Guard.NotNull(build, nameof(build));
        var nestedCanvas = new PdfPageCanvas(allowOutOfPageCoordinates: true);
        build(nestedCanvas);
        if (nestedCanvas.Items.Count == 0) {
            throw new ArgumentException("Canvas figures require at least one content item.", nameof(build));
        }
        _items.Add(new PdfCanvasFigureItem(alternativeText.Trim(), nestedCanvas.Items));
        return this;
    }

    /// <summary>
    /// Groups positioned text fragments under one logical replacement string for extraction and accessibility.
    /// Child paint remains unchanged while readers that honor <c>ActualText</c> receive the supplied logical text once.
    /// </summary>
    public PdfPageCanvas ActualText(string text, Action<PdfPageCanvas> build) {
        Guard.NotNull(text, nameof(text));
        if (text.Length == 0) throw new ArgumentException("Canvas actual text cannot be empty.", nameof(text));
        Guard.NotNull(build, nameof(build));
        var nestedCanvas = new PdfPageCanvas(allowOutOfPageCoordinates: true);
        build(nestedCanvas);
        if (nestedCanvas.Items.Count == 0) {
            throw new ArgumentException("Canvas actual-text groups require at least one content item.", nameof(build));
        }
        _items.Add(new PdfCanvasActualTextItem(text, nestedCanvas.Items));
        return this;
    }

    /// <summary>Groups absolute canvas content under a typed tagged-PDF structure container.</summary>
    public PdfPageCanvas Structure(PdfCanvasStructureRole role, Action<PdfPageCanvas> build, PdfCanvasStructureOptions? options = null) {
        if ((int)role < (int)PdfCanvasStructureRole.Section || (int)role > (int)PdfCanvasStructureRole.Caption) {
            throw new ArgumentOutOfRangeException(nameof(role));
        }
        Guard.NotNull(build, nameof(build));
        PdfCanvasStructureOptions snapshot = options?.Clone() ?? new PdfCanvasStructureOptions();
        bool tableCell = role == PdfCanvasStructureRole.TableHeaderCell || role == PdfCanvasStructureRole.TableCell;
        if (!tableCell && (snapshot.ColumnSpan != 1 || snapshot.RowSpan != 1)) {
            throw new ArgumentException("Canvas structure spans are valid only for table cells.", nameof(options));
        }
        if (role != PdfCanvasStructureRole.TableHeaderCell && snapshot.HeaderScope.HasValue) {
            throw new ArgumentException("Canvas table header scope is valid only for table-header cells.", nameof(options));
        }

        var nestedCanvas = new PdfPageCanvas(allowOutOfPageCoordinates: true);
        build(nestedCanvas);
        if (nestedCanvas.Items.Count == 0) {
            throw new ArgumentException("Canvas structure containers require at least one content item.", nameof(build));
        }
        _items.Add(new PdfCanvasStructureItem(role, snapshot, nestedCanvas.Items));
        return this;
    }

    /// <summary>Adds text inside a fixed page rectangle using top-left page coordinates.</summary>
    public PdfPageCanvas Text(string text, double x, double y, double width, double height, double? fontSize = null, PdfColor? color = null, PdfAlign align = PdfAlign.Left, PdfStandardFont? font = null) {
        Guard.NotNull(text, nameof(text));
        return Text(new[] { TextRun.Normal(text, color, fontSize, font: font) }, x, y, width, height, color, align, fontSize);
    }

    /// <summary>Adds rich text runs inside a fixed page rectangle using top-left page coordinates.</summary>
    public PdfPageCanvas Text(IEnumerable<TextRun> runs, double x, double y, double width, double height, PdfColor? defaultColor = null, PdfAlign align = PdfAlign.Left, double? fontSize = null, double? lineHeight = null) {
        return AddText(runs, PdfCanvasTextStructureRole.Paragraph, x, y, width, height, defaultColor, align, fontSize, lineHeight);
    }

    /// <summary>Adds tagged rich text runs inside a fixed page rectangle using top-left page coordinates.</summary>
    public PdfPageCanvas Text(IEnumerable<TextRun> runs, PdfCanvasTextStructureRole structureRole, double x, double y, double width, double height, PdfColor? defaultColor = null, PdfAlign align = PdfAlign.Left, double? fontSize = null, double? lineHeight = null) {
        return AddText(runs, structureRole, x, y, width, height, defaultColor, align, fontSize, lineHeight);
    }

    private PdfPageCanvas AddText(IEnumerable<TextRun> runs, PdfCanvasTextStructureRole structureRole, double x, double y, double width, double height, PdfColor? defaultColor, PdfAlign align, double? fontSize, double? lineHeight) {
        Guard.NotNull(runs, nameof(runs));
        if ((int)structureRole < (int)PdfCanvasTextStructureRole.Paragraph
            || (int)structureRole > (int)PdfCanvasTextStructureRole.Span) {
            throw new ArgumentOutOfRangeException(nameof(structureRole));
        }
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
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

        _items.Add(new PdfCanvasTextItem(snapshot, x, y, width, height, defaultColor, align, fontSize, lineHeight, structureRole));
        return this;
    }

    /// <summary>Adds a styled fixed-position text box using top-left page coordinates.</summary>
    public PdfPageCanvas TextBox(string text, double x, double y, double width, double height, PdfCanvasTextBoxStyle? style = null, double rotationAngle = 0D) {
        return TextBox(text, x, y, width, height, style, rotationAngle, diagnosticHandler: null);
    }

    /// <summary>Adds a styled fixed-position text box using top-left page coordinates and reports layout diagnostics during rendering.</summary>
    public PdfPageCanvas TextBox(string text, double x, double y, double width, double height, PdfCanvasTextBoxStyle? style, Action<PdfLayoutDiagnostic>? diagnosticHandler) {
        return TextBox(text, x, y, width, height, style, 0D, diagnosticHandler);
    }

    /// <summary>Adds a styled fixed-position text box using top-left page coordinates and reports layout diagnostics during rendering.</summary>
    public PdfPageCanvas TextBox(string text, double x, double y, double width, double height, PdfCanvasTextBoxStyle? style, double rotationAngle, Action<PdfLayoutDiagnostic>? diagnosticHandler) {
        Guard.NotNull(text, nameof(text));
        return TextBox(new[] { TextRun.Normal(text) }, x, y, width, height, style, rotationAngle, diagnosticHandler);
    }

    /// <summary>Adds styled rich text inside a fixed-position text box using top-left page coordinates.</summary>
    public PdfPageCanvas TextBox(IEnumerable<TextRun> runs, double x, double y, double width, double height, PdfCanvasTextBoxStyle? style = null, double rotationAngle = 0D) {
        return TextBox(runs, x, y, width, height, style, rotationAngle, diagnosticHandler: null);
    }

    /// <summary>Adds styled rich text inside a fixed-position text box using top-left page coordinates and reports layout diagnostics during rendering.</summary>
    public PdfPageCanvas TextBox(IEnumerable<TextRun> runs, double x, double y, double width, double height, PdfCanvasTextBoxStyle? style, Action<PdfLayoutDiagnostic>? diagnosticHandler) {
        return TextBox(runs, x, y, width, height, style, 0D, diagnosticHandler);
    }

    /// <summary>Adds styled rich text inside a fixed-position text box using top-left page coordinates and reports layout diagnostics during rendering.</summary>
    public PdfPageCanvas TextBox(IEnumerable<TextRun> runs, double x, double y, double width, double height, PdfCanvasTextBoxStyle? style, double rotationAngle, Action<PdfLayoutDiagnostic>? diagnosticHandler) {
        Guard.NotNull(runs, nameof(runs));
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        ValidateFiniteRotation(rotationAngle, nameof(rotationAngle), "Canvas text box rotation angle must be finite.");

        PdfCanvasTextBoxStyle textBoxStyle = style?.Clone() ?? new PdfCanvasTextBoxStyle();
        ValidateTextBoxInnerArea(width, height, textBoxStyle, nameof(style));

        var snapshot = runs.ToList();
        if (snapshot.Count == 0) {
            throw new ArgumentException("Canvas text box requires at least one text run.", nameof(runs));
        }

        _items.Add(new PdfCanvasTextBoxItem(ApplyTextBoxDefaults(snapshot, textBoxStyle), x, y, width, height, textBoxStyle, rotationAngle, diagnosticHandler));
        return this;
    }

    /// <summary>Adds a shared drawing shape at fixed top-left page coordinates.</summary>
    public PdfPageCanvas Shape(OfficeShape shape, double x, double y, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null, double rotationAngle = 0D) {
        Guard.NotNull(shape, nameof(shape));
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
        ValidateFiniteRotation(rotationAngle, nameof(rotationAngle), "Canvas shape rotation angle must be finite.");
        ShapeBlock block = PdfDocument.CreateShapeBlock(CreateRotatedShape(shape, rotationAngle), PdfAlign.Left, spacingBefore: 0D, spacingAfter: 0D, style, linkUri, linkContents);
        _items.Add(new PdfCanvasShapeItem(block, x, y, rotationAngle));
        return this;
    }

    /// <summary>Adds a shared vector drawing inside a fixed page rectangle using top-left page coordinates.</summary>
    public PdfPageCanvas Drawing(OfficeDrawing drawing, double x, double y, double width, double height, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null, double rotationAngle = 0D) {
        Guard.NotNull(drawing, nameof(drawing));
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
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
        return ImageShared(
            PdfCanvasImageResource.Create(imageBytes),
            x,
            y,
            width,
            height,
            style,
            linkUri,
            linkContents,
            alternativeText,
            rotationAngle,
            horizontalFlip,
            verticalFlip);
    }

    internal PdfPageCanvas ImageShared(PdfCanvasImageResource imageResource, double x, double y, double width, double height, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null, string? alternativeText = null, double rotationAngle = 0D, bool horizontalFlip = false, bool verticalFlip = false) {
        Guard.NotNull(imageResource, nameof(imageResource));
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.OptionalUriAction(linkUri, nameof(linkUri));
        ValidateFiniteRotation(rotationAngle, nameof(rotationAngle), "Canvas image rotation angle must be finite.");

        PdfImageStyle? imageStyle = PdfDocument.CreateImageStyle(PdfAlign.Left, style?.ClipPath, style?.Fit, spacingBefore: 0D, spacingAfter: 0D, style, alternativeText);
        if (imageStyle != null) {
            PdfDocument.ValidateImageStyleForBox(imageStyle, width, height, nameof(style));
        }

        OfficeImageInfo imageInfo = imageResource.Info;
        if (imageStyle != null) {
            PdfDocument.ValidateImageFitDimensions(imageInfo, imageStyle.Fit, nameof(style));
        }

        _items.Add(new PdfCanvasImageItem(new ImageBlock(imageResource.Bytes, width, height, imageInfo, imageStyle, linkUri, linkContents, useDataSnapshot: true), x, y, rotationAngle, horizontalFlip, verticalFlip));
        return this;
    }

    /// <summary>Adds a PDF text annotation at fixed top-left page coordinates.</summary>
    public PdfPageCanvas TextAnnotation(string contents, double x, double y, double width = 18D, double height = 18D, PdfTextAnnotationIcon icon = PdfTextAnnotationIcon.Comment, PdfColor? color = null, bool open = false) {
        Guard.NotNullOrWhiteSpace(contents, nameof(contents));
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        PdfDocument.ValidateTextAnnotationIcon(icon, nameof(icon));
        _items.Add(new PdfCanvasTextAnnotationItem(contents, x, y, width, height, icon, color, open));
        return this;
    }

    /// <summary>Adds a PDF free-text annotation at fixed top-left page coordinates.</summary>
    public PdfPageCanvas FreeTextAnnotation(string contents, double x, double y, double width, double height, double fontSize = 10D, PdfColor? textColor = null, PdfColor? borderColor = null, double borderWidth = 1D, PdfColor? fillColor = null, PdfAlign textAlign = PdfAlign.Left, double padding = 3D, double? lineHeight = null) {
        Guard.NotNullOrWhiteSpace(contents, nameof(contents));
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.Positive(fontSize, nameof(fontSize));
        Guard.NonNegative(borderWidth, nameof(borderWidth));
        Guard.LeftCenterRightAlign(textAlign, nameof(textAlign), "Canvas free text annotation text");
        Guard.NonNegative(padding, nameof(padding));
        if (lineHeight.HasValue) {
            Guard.Positive(lineHeight.Value, nameof(lineHeight));
        }

        _items.Add(new PdfCanvasFreeTextAnnotationItem(contents, x, y, width, height, fontSize, textColor ?? PdfColor.Black, borderColor, borderWidth, fillColor, textAlign, padding, lineHeight));
        return this;
    }

    /// <summary>Adds a PDF highlight annotation rectangle at fixed top-left page coordinates.</summary>
    public PdfPageCanvas HighlightAnnotation(string contents, double x, double y, double width, double height, PdfColor? color = null) {
        Guard.NotNullOrWhiteSpace(contents, nameof(contents));
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        _items.Add(new PdfCanvasHighlightAnnotationItem(contents, x, y, width, height, color ?? new PdfColor(1D, 0.92D, 0.2D)));
        return this;
    }

    /// <summary>Adds a fixed-position table inside a page rectangle using top-left page coordinates.</summary>
    public PdfPageCanvas Table(IEnumerable<string[]> rows, double x, double y, double width, double height, PdfTableStyle? style = null, double rotationAngle = 0D) {
        return Table(rows, x, y, width, height, style, rotationAngle, diagnosticHandler: null);
    }

    /// <summary>Adds a fixed-position table inside a page rectangle using top-left page coordinates and reports layout diagnostics during rendering.</summary>
    public PdfPageCanvas Table(IEnumerable<string[]> rows, double x, double y, double width, double height, PdfTableStyle? style, Action<PdfLayoutDiagnostic>? diagnosticHandler) {
        return Table(rows, x, y, width, height, style, 0D, diagnosticHandler);
    }

    /// <summary>Adds a fixed-position table inside a page rectangle using top-left page coordinates and reports layout diagnostics during rendering.</summary>
    public PdfPageCanvas Table(IEnumerable<string[]> rows, double x, double y, double width, double height, PdfTableStyle? style, double rotationAngle, Action<PdfLayoutDiagnostic>? diagnosticHandler) {
        Guard.NotNull(rows, nameof(rows));
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        ValidateFiniteRotation(rotationAngle, nameof(rotationAngle), "Canvas table rotation angle must be finite.");
        _items.Add(new PdfCanvasTableItem(new TableBlock(rows, PdfAlign.Left, style), x, y, width, height, rotationAngle, diagnosticHandler));
        return this;
    }

    /// <summary>Adds a fixed-position rich table inside a page rectangle using top-left page coordinates.</summary>
    public PdfPageCanvas Table(IEnumerable<PdfTableCell[]> rows, double x, double y, double width, double height, PdfTableStyle? style = null, double rotationAngle = 0D) {
        return Table(rows, x, y, width, height, style, rotationAngle, diagnosticHandler: null);
    }

    /// <summary>Adds a fixed-position rich table inside a page rectangle using top-left page coordinates and reports layout diagnostics during rendering.</summary>
    public PdfPageCanvas Table(IEnumerable<PdfTableCell[]> rows, double x, double y, double width, double height, PdfTableStyle? style, Action<PdfLayoutDiagnostic>? diagnosticHandler) {
        return Table(rows, x, y, width, height, style, 0D, diagnosticHandler);
    }

    /// <summary>Adds a fixed-position rich table inside a page rectangle using top-left page coordinates and reports layout diagnostics during rendering.</summary>
    public PdfPageCanvas Table(IEnumerable<PdfTableCell[]> rows, double x, double y, double width, double height, PdfTableStyle? style, double rotationAngle, Action<PdfLayoutDiagnostic>? diagnosticHandler) {
        Guard.NotNull(rows, nameof(rows));
        ValidateCanvasCoordinate(x, nameof(x));
        ValidateCanvasCoordinate(y, nameof(y));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        ValidateFiniteRotation(rotationAngle, nameof(rotationAngle), "Canvas table rotation angle must be finite.");
        _items.Add(new PdfCanvasTableItem(new TableBlock(rows, PdfAlign.Left, style), x, y, width, height, rotationAngle, diagnosticHandler));
        return this;
    }

    /// <summary>Adds a clipped fixed-position canvas frame using top-left page coordinates.</summary>
    public PdfPageCanvas Clip(double x, double y, double width, double height, Action<PdfPageCanvas> build) =>
        Clip(x, y, OfficeClipPath.Rectangle(width, height), build);

    /// <summary>Adds a path-clipped fixed-position canvas frame using top-left page coordinates.</summary>
    public PdfPageCanvas Clip(double x, double y, OfficeClipPath clipPath, Action<PdfPageCanvas> build) {
        Guard.NonNegative(x, nameof(x));
        Guard.NonNegative(y, nameof(y));
        Guard.NotNull(clipPath, nameof(clipPath));
        Guard.NotNull(build, nameof(build));

        var clippedCanvas = new PdfPageCanvas(allowOutOfPageCoordinates: true);
        build(clippedCanvas);
        _items.Add(new PdfCanvasClipItem(clippedCanvas.Items, x, y, clipPath));
        return this;
    }

    /// <summary>Adds nested canvas content through one top-left-coordinate affine transform and opacity state.</summary>
    public PdfPageCanvas Effect(OfficeTransform transform, double opacity, Action<PdfPageCanvas> build) {
        if (double.IsNaN(opacity) || double.IsInfinity(opacity) || opacity < 0D || opacity > 1D) {
            throw new ArgumentOutOfRangeException(nameof(opacity), "Canvas effect opacity must be between zero and one.");
        }
        Guard.NotNull(build, nameof(build));
        var nestedCanvas = new PdfPageCanvas(allowOutOfPageCoordinates: true);
        build(nestedCanvas);
        _items.Add(new PdfCanvasEffectItem(nestedCanvas.Items, transform, opacity));
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
                run.BackgroundColor,
                run.FontFamily));
        }

        return styled.AsReadOnly();
    }

    private static void ValidateTextBoxInnerArea(double width, double height, PdfCanvasTextBoxStyle style, string paramName) {
        if (style.EffectivePaddingLeft + style.EffectivePaddingRight >= width ||
            style.EffectivePaddingTop + style.EffectivePaddingBottom >= height) {
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

    private void ValidateCanvasCoordinate(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Canvas coordinates must be finite.");
        }

        if (!_allowOutOfPageCoordinates) {
            Guard.NonNegative(value, paramName);
        }
    }
}

internal sealed class PdfCanvasImageResource {
    private PdfCanvasImageResource(byte[] bytes, OfficeImageInfo info) {
        Bytes = bytes;
        Info = info;
    }

    internal byte[] Bytes { get; }
    internal OfficeImageInfo Info { get; }

    internal static PdfCanvasImageResource Create(byte[] bytes) {
        Guard.NotNullOrEmpty(bytes, nameof(bytes));
        byte[] snapshot = (byte[])bytes.Clone();
        return new PdfCanvasImageResource(snapshot, PdfDocument.ValidateImageBytes(snapshot));
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

internal sealed class PdfCanvasOutlineItem : PdfCanvasItem {
    public PdfCanvasOutlineItem(string title, int level, double y)
        : base(0D, y) {
        Title = title;
        Level = level;
    }

    public string Title { get; }
    public int Level { get; }
}

internal sealed class PdfCanvasFigureItem : PdfCanvasItem {
    public PdfCanvasFigureItem(string alternativeText, IReadOnlyList<PdfCanvasItem> items)
        : base(0D, 0D) {
        AlternativeText = alternativeText;
        Items = items;
    }

    public string AlternativeText { get; }
    public IReadOnlyList<PdfCanvasItem> Items { get; }
}

internal sealed class PdfCanvasStructureItem : PdfCanvasItem {
    public PdfCanvasStructureItem(PdfCanvasStructureRole role, PdfCanvasStructureOptions options, IReadOnlyList<PdfCanvasItem> items)
        : base(0D, 0D) {
        Role = role;
        Options = options;
        Items = items;
    }

    public PdfCanvasStructureRole Role { get; }
    public PdfCanvasStructureOptions Options { get; }
    public IReadOnlyList<PdfCanvasItem> Items { get; }
}

internal sealed class PdfCanvasActualTextItem : PdfCanvasItem {
    public PdfCanvasActualTextItem(string text, IReadOnlyList<PdfCanvasItem> items)
        : base(0D, 0D) {
        Text = text;
        Items = items;
    }

    public string Text { get; }
    public IReadOnlyList<PdfCanvasItem> Items { get; }
}

internal sealed class PdfCanvasTextItem : PdfCanvasItem {
    public PdfCanvasTextItem(IReadOnlyList<TextRun> runs, double x, double y, double width, double height, PdfColor? defaultColor, PdfAlign align, double? fontSize, double? lineHeight, PdfCanvasTextStructureRole structureRole)
        : base(x, y) {
        Runs = runs;
        Width = width;
        Height = height;
        DefaultColor = defaultColor;
        Align = align;
        FontSize = fontSize;
        LineHeight = lineHeight;
        StructureRole = structureRole;
    }

    public IReadOnlyList<TextRun> Runs { get; }
    public double Width { get; }
    public double Height { get; }
    public PdfColor? DefaultColor { get; }
    public PdfAlign Align { get; }
    public double? FontSize { get; }
    public double? LineHeight { get; }
    public PdfCanvasTextStructureRole StructureRole { get; }
}

internal sealed class PdfCanvasTextBoxItem : PdfCanvasItem {
    public PdfCanvasTextBoxItem(IReadOnlyList<TextRun> runs, double x, double y, double width, double height, PdfCanvasTextBoxStyle style, double rotationAngle, Action<PdfLayoutDiagnostic>? diagnosticHandler)
        : base(x, y) {
        Runs = runs;
        Width = width;
        Height = height;
        Style = style.Clone();
        RotationAngle = rotationAngle;
        DiagnosticHandler = diagnosticHandler;
    }

    public IReadOnlyList<TextRun> Runs { get; }
    public double Width { get; }
    public double Height { get; }
    public PdfCanvasTextBoxStyle Style { get; }
    public double RotationAngle { get; }
    public Action<PdfLayoutDiagnostic>? DiagnosticHandler { get; }
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

internal sealed class PdfCanvasTextAnnotationItem : PdfCanvasItem {
    public PdfCanvasTextAnnotationItem(string contents, double x, double y, double width, double height, PdfTextAnnotationIcon icon, PdfColor? color, bool open)
        : base(x, y) {
        Contents = contents;
        Width = width;
        Height = height;
        Icon = icon;
        Color = color;
        Open = open;
    }

    public string Contents { get; }
    public double Width { get; }
    public double Height { get; }
    public PdfTextAnnotationIcon Icon { get; }
    public PdfColor? Color { get; }
    public bool Open { get; }
}

internal sealed class PdfCanvasFreeTextAnnotationItem : PdfCanvasItem {
    public PdfCanvasFreeTextAnnotationItem(string contents, double x, double y, double width, double height, double fontSize, PdfColor textColor, PdfColor? borderColor, double borderWidth, PdfColor? fillColor, PdfAlign textAlign, double padding, double? lineHeight)
        : base(x, y) {
        Contents = contents;
        Width = width;
        Height = height;
        FontSize = fontSize;
        TextColor = textColor;
        BorderColor = borderColor;
        BorderWidth = borderWidth;
        FillColor = fillColor;
        TextAlign = textAlign;
        Padding = padding;
        LineHeight = lineHeight;
    }

    public string Contents { get; }
    public double Width { get; }
    public double Height { get; }
    public double FontSize { get; }
    public PdfColor TextColor { get; }
    public PdfColor? BorderColor { get; }
    public double BorderWidth { get; }
    public PdfColor? FillColor { get; }
    public PdfAlign TextAlign { get; }
    public double Padding { get; }
    public double? LineHeight { get; }
}

internal sealed class PdfCanvasHighlightAnnotationItem : PdfCanvasItem {
    public PdfCanvasHighlightAnnotationItem(string contents, double x, double y, double width, double height, PdfColor color)
        : base(x, y) {
        Contents = contents;
        Width = width;
        Height = height;
        Color = color;
    }

    public string Contents { get; }
    public double Width { get; }
    public double Height { get; }
    public PdfColor Color { get; }
}

internal sealed class PdfCanvasTableItem : PdfCanvasItem {
    public PdfCanvasTableItem(TableBlock block, double x, double y, double width, double height, double rotationAngle, Action<PdfLayoutDiagnostic>? diagnosticHandler)
        : base(x, y) {
        Block = block;
        Width = width;
        Height = height;
        RotationAngle = rotationAngle;
        DiagnosticHandler = diagnosticHandler;
    }

    public TableBlock Block { get; }
    public double Width { get; }
    public double Height { get; }
    public double RotationAngle { get; }
    public Action<PdfLayoutDiagnostic>? DiagnosticHandler { get; }
}

internal sealed class PdfCanvasClipItem : PdfCanvasItem {
    public PdfCanvasClipItem(IReadOnlyList<PdfCanvasItem> items, double x, double y, OfficeClipPath clipPath)
        : base(x, y) {
        Items = items;
        ClipPath = clipPath.Clone();
    }

    public IReadOnlyList<PdfCanvasItem> Items { get; }
    public OfficeClipPath ClipPath { get; }
    public double Width => ClipPath.Width;
    public double Height => ClipPath.Height;
}

internal sealed class PdfCanvasEffectItem : PdfCanvasItem {
    public PdfCanvasEffectItem(IReadOnlyList<PdfCanvasItem> items, OfficeTransform transform, double opacity)
        : base(0D, 0D) {
        Items = items;
        Transform = transform;
        Opacity = opacity;
    }

    public IReadOnlyList<PdfCanvasItem> Items { get; }
    public OfficeTransform Transform { get; }
    public double Opacity { get; }
}
