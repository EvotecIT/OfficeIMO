using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDoc {
    /// <summary>Adds a level-1 heading.</summary>
    public PdfDoc H1(string text, PdfAlign align = PdfAlign.Left, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) {
        Guard.NotNullOrWhiteSpace(text, nameof(text));
        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
        AddBlock(new HeadingBlock(1, text, align, color, linkUri, style, linkContents, linkDestinationName)); return this;
    }
    /// <summary>Adds a level-2 heading.</summary>
    public PdfDoc H2(string text, PdfAlign align = PdfAlign.Left, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) {
        Guard.NotNullOrWhiteSpace(text, nameof(text));
        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
        AddBlock(new HeadingBlock(2, text, align, color, linkUri, style, linkContents, linkDestinationName)); return this;
    }
    /// <summary>Adds a level-3 heading.</summary>
    public PdfDoc H3(string text, PdfAlign align = PdfAlign.Left, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) {
        Guard.NotNullOrWhiteSpace(text, nameof(text));
        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
        AddBlock(new HeadingBlock(3, text, align, color, linkUri, style, linkContents, linkDestinationName)); return this;
    }

    /// <summary>Inserts a page break.</summary>
    public PdfDoc PageBreak() { AddBlock(new PageBreakBlock()); return this; }

    /// <summary>Adds invisible vertical space to the current document flow.</summary>
    public PdfDoc Spacer(double height) {
        AddBlock(new SpacerBlock(height));
        return this;
    }

    /// <summary>Configures a page-scoped flow with its own page setup and default styles.</summary>
    public PdfDoc Page(System.Action<PdfPageCompose> configure) {
        AddComposedPage(configure);
        return this;
    }

    /// <summary>Configures a section-scoped flow with its own page setup and default styles.</summary>
    public PdfDoc Section(System.Action<PdfPageCompose> configure) {
        AddComposedPage(configure);
        return this;
    }

    /// <summary>Sets the document-wide default page size used by top-level flow and composed pages.</summary>
    public PdfDoc Size(PageSize size) {
        _options.PageSize = size;
        return this;
    }

    /// <summary>Sets the document-wide default page size in points.</summary>
    public PdfDoc Size(double width, double height) {
        _options.PageSize = new PageSize(width, height);
        return this;
    }

    /// <summary>Sets the document-wide default page orientation while preserving the current page size dimensions.</summary>
    public PdfDoc Orientation(PdfPageOrientation orientation) {
        _options.PageSize = _options.PageSize.WithOrientation(orientation);
        return this;
    }

    /// <summary>Sets or clears the document-wide default page background color.</summary>
    public PdfDoc Background(PdfColor? color) {
        _options.BackgroundColor = color;
        return this;
    }

    /// <summary>Sets the document-wide default page orientation to portrait.</summary>
    public PdfDoc Portrait() => Orientation(PdfPageOrientation.Portrait);

    /// <summary>Sets the document-wide default page orientation to landscape.</summary>
    public PdfDoc Landscape() => Orientation(PdfPageOrientation.Landscape);

    /// <summary>Sets uniform document-wide default page margins in points.</summary>
    public PdfDoc Margin(double all) {
        _options.Margins = PageMargins.Uniform(all);
        return this;
    }

    /// <summary>Sets document-wide default page margins from a reusable margin value.</summary>
    public PdfDoc Margin(PageMargins margins) {
        _options.Margins = margins;
        return this;
    }

    /// <summary>Sets document-wide default page margins in points.</summary>
    public PdfDoc Margin(double left, double top, double right, double bottom) {
        _options.Margins = new PageMargins(left, top, right, bottom);
        return this;
    }

    /// <summary>Sets the first visible page number for the document-wide flow.</summary>
    public PdfDoc PageNumberStart(int start) {
        _options.PageNumberStart = start;
        return this;
    }

    /// <summary>Sets the document-wide visible page-number style for header/footer tokens.</summary>
    public PdfDoc PageNumberStyle(PdfPageNumberStyle style) {
        _options.PageNumberStyle = style;
        return this;
    }

    /// <summary>Defines the document-wide default header layout and content.</summary>
    public PdfDoc Header(System.Action<PdfHeaderCompose> build) {
        Guard.NotNull(build, nameof(build));
        var header = new PdfHeaderCompose(_options);
        build(header);
        return this;
    }

    /// <summary>Defines the document-wide default footer layout and content.</summary>
    public PdfDoc Footer(System.Action<PdfFooterCompose> build) {
        Guard.NotNull(build, nameof(build));
        var footer = new PdfFooterCompose(_options);
        build(footer);
        return this;
    }

    /// <summary>Applies reusable document-wide default styles.</summary>
    public PdfDoc Theme(PdfTheme theme) {
        Guard.NotNull(theme, nameof(theme));
        theme.Clone().ApplyTo(_options);
        return this;
    }

    /// <summary>Sets document-wide default text styling used by following page-flow content.</summary>
    public PdfDoc DefaultTextStyle(System.Action<PdfTextStyleCompose> style) {
        Guard.NotNull(style, nameof(style));
        var compose = new PdfTextStyleCompose(_options);
        style(compose);
        return this;
    }

    /// <summary>Sets document-wide default text styling from a reusable text style object.</summary>
    public PdfDoc DefaultTextStyle(PdfTextStyle style) {
        Guard.NotNull(style, nameof(style));
        style.Clone().ApplyTo(_options);
        return this;
    }

    /// <summary>Sets the document-wide default paragraph style used by paragraphs that do not provide an explicit style.</summary>
    public PdfDoc DefaultParagraphStyle(PdfParagraphStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultParagraphStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default table style used by tables that do not provide an explicit style.</summary>
    public PdfDoc DefaultTableStyle(PdfTableStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultTableStyle = style;
        return this;
    }

    /// <summary>Sets the document-wide default table style from a supported Word table style name.</summary>
    public PdfDoc DefaultTableStyle(string wordTableStyleName) {
        _options.DefaultTableStyle = TableStyles.FromWordTableStyle(wordTableStyleName);
        return this;
    }

    /// <summary>Sets the document-wide default style for a built-in heading level.</summary>
    public PdfDoc DefaultHeadingStyle(int level, PdfHeadingStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.SetDefaultHeadingStyle(level, style);
        return this;
    }

    /// <summary>Adds a simple bullet list.</summary>
    public PdfDoc Bullets(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        AddBlock(new BulletListBlock(items, align, color, style));
        return this;
    }

    /// <summary>Adds a bullet list whose items can contain rich inline text runs.</summary>
    public PdfDoc RichBullets(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        AddBlock(new BulletListBlock(items, align, color, style));
        return this;
    }

    /// <summary>Adds a simple numbered list.</summary>
    public PdfDoc Numbered(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, int startNumber = 1, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        AddBlock(new NumberedListBlock(items, align, color, startNumber, style));
        return this;
    }

    /// <summary>Adds a numbered list whose items can contain rich inline text runs.</summary>
    public PdfDoc RichNumbered(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, int startNumber = 1, PdfListStyle? style = null) {
        Guard.NotNull(items, nameof(items));
        AddBlock(new NumberedListBlock(items, align, color, startNumber, style));
        return this;
    }

    /// <summary>Sets the document-wide default style for bullet and numbered lists.</summary>
    public PdfDoc DefaultListStyle(PdfListStyle style) {
        Guard.NotNull(style, nameof(style));
        _options.DefaultListStyle = style;
        return this;
    }

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

    /// <summary>
    /// Fluent paragraph builder in the style of OfficeIMO.Word.
    /// Example: Paragraph(p => p.Color(PdfColor.Gray).Text("You can ").Bold("mix ").Italic("styles"));
    /// </summary>
    public PdfDoc Paragraph(System.Action<PdfParagraphBuilder> compose, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null, PdfParagraphStyle? style = null) {
        Guard.NotNull(compose, nameof(compose));
        var builder = new PdfParagraphBuilder(align, defaultColor);
        compose(builder);
        AddBlock(builder.Build(style));
        return this;
    }

    /// <summary>
    /// Higher-level composition model (page size/margins/footer + content), similar to other document DSLs.
    /// Sugar only; composes into the same PdfDoc blocks and options.
    /// </summary>
    public PdfDoc Compose(System.Action<PdfCompose> compose) {
        Guard.NotNull(compose, nameof(compose));
        var c = new PdfCompose(this);
        compose(c);
        return this;
    }

    /// <summary>Adds a horizontal rule (line) spanning the content width.</summary>
    public PdfDoc HR(double? thickness = null, PdfColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfHorizontalRuleStyle? style = null) {
        AddBlock(new HorizontalRuleBlock(CreateHorizontalRuleStyle(thickness, color, spacingBefore, spacingAfter, style)));
        return this;
    }

    /// <summary>Adds a named bookmark at the current flow position.</summary>
    public PdfDoc Bookmark(string name) {
        AddBlock(new BookmarkBlock(name));
        return this;
    }

    /// <summary>Adds a simple AcroForm text field at the current flow position.</summary>
    public PdfDoc TextField(string name, double width = 180, double height = 22, string value = "", PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) {
        AddBlock(new TextFieldBlock(name, width, height, value, align, fontSize, spacingBefore, spacingAfter, style));
        return this;
    }

    /// <summary>Adds a simple AcroForm check box at the current flow position.</summary>
    public PdfDoc CheckBox(string name, bool isChecked = false, double size = 14, PdfAlign align = PdfAlign.Left, double spacingBefore = 0, double spacingAfter = 6, string checkedValueName = "Yes", PdfFormFieldStyle? style = null) {
        AddBlock(new CheckBoxBlock(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style));
        return this;
    }

    /// <summary>Adds a simple AcroForm choice field at the current flow position.</summary>
    public PdfDoc ChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, double width = 180, double height = 22, PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, bool isComboBox = true, PdfFormFieldStyle? style = null) {
        AddBlock(new ChoiceFieldBlock(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style));
        return this;
    }

    /// <summary>Adds a simple AcroForm multi-select choice field at the current flow position.</summary>
    public PdfDoc MultiSelectChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, System.Collections.Generic.IEnumerable<string>? values = null, double width = 180, double height = 72, PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) {
        AddBlock(new ChoiceFieldBlock(name, options, values, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox: false, allowsMultipleSelection: true, style));
        return this;
    }

    /// <summary>Adds a simple AcroForm radio button group at the current flow position.</summary>
    public PdfDoc RadioButtonGroup(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, double size = 14, double gap = 6, PdfAlign align = PdfAlign.Left, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) {
        AddBlock(new RadioButtonGroupBlock(name, options, value, size, gap, align, spacingBefore, spacingAfter, style));
        return this;
    }

    /// <summary>Adds a shared OfficeIMO.Drawing shape at the current flow position.</summary>
    public PdfDoc Shape(OfficeShape shape, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        AddBlock(CreateShapeBlock(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents));
        return this;
    }

    /// <summary>Adds a shared OfficeIMO.Drawing scene at the current flow position.</summary>
    public PdfDoc Drawing(OfficeDrawing drawing, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        AddBlock(CreateDrawingBlock(drawing, align, spacingBefore, spacingAfter, style, linkUri, linkContents));
        return this;
    }

    /// <summary>Adds a flow line using the shared OfficeIMO.Drawing shape descriptor.</summary>
    public PdfDoc Line(double x1, double y1, double x2, double y2, PdfColor? strokeColor = null, double strokeWidth = 1, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Line(x1, y1, x2, y2);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        return Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a flow rectangle using the shared OfficeIMO.Drawing shape descriptor.</summary>
    public PdfDoc Rectangle(double width, double height, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Rectangle(width, height);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        return Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a flow rounded rectangle using the shared OfficeIMO.Drawing shape descriptor.</summary>
    public PdfDoc RoundedRectangle(double width, double height, double cornerRadius, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.RoundedRectangle(width, height, cornerRadius);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        return Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a flow ellipse using the shared OfficeIMO.Drawing shape descriptor.</summary>
    public PdfDoc Ellipse(double width, double height, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Ellipse(width, height);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        return Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a flow polygon using the shared OfficeIMO.Drawing shape descriptor.</summary>
    public PdfDoc Polygon(System.Collections.Generic.IEnumerable<OfficePoint> points, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Polygon(points);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        return Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a flow path using the shared OfficeIMO.Drawing shape descriptor.</summary>
    public PdfDoc Path(System.Collections.Generic.IEnumerable<OfficePathCommand> commands, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Path(commands);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        return Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a paragraph inside a simple panel (background + optional border).</summary>
    public PdfDoc PanelParagraph(System.Action<PdfParagraphBuilder> compose, PanelStyle? style = null, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) {
        Guard.NotNull(compose, nameof(compose));
        Guard.ParagraphAlign(align, nameof(align), "Panel paragraph");
        var builder = new PdfParagraphBuilder(align, defaultColor);
        compose(builder);
        AddBlock(new PanelParagraphBlock(builder.Build().Runs, align, defaultColor, style));
        return this;
    }

    /// <summary>Adds a supported image at the current flow position. JPEG and simple non-interlaced 8-bit PNG images, including grayscale-alpha/RGBA soft masks, are currently supported.</summary>
    public PdfDoc Image(byte[] jpegBytes, double width, double height, PdfAlign? align = null, OfficeClipPath? clipPath = null, OfficeImageFit? fit = null, double? spacingBefore = null, double? spacingAfter = null, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null) {
        Guard.NotNullOrEmpty(jpegBytes, nameof(jpegBytes));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
        PdfImageStyle? imageStyle = CreateImageStyle(align, clipPath, fit, spacingBefore, spacingAfter, style);
        if (imageStyle != null) {
            ValidateImageStyleForBox(imageStyle, width, height, nameof(clipPath));
        }

        var imageInfo = ValidateImageBytes(jpegBytes);
        if (imageStyle != null) {
            ValidateImageFitDimensions(imageInfo, imageStyle.Fit, nameof(fit));
        }

        AddBlock(new ImageBlock(jpegBytes, width, height, imageInfo, imageStyle, linkUri, linkContents));
        return this;
    }

    // Internal for Compose Row
    internal void AddRow(RowBlock row) { AddBlock(row); }

    /// <summary>Adds a simple table from rows of string arrays.</summary>
    public PdfDoc Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        AddBlock(new TableBlock(rows, align, style));
        return this;
    }

    /// <summary>Adds a table from explicit cells, including optional column spans.</summary>
    public PdfDoc Table(System.Collections.Generic.IEnumerable<PdfTableCell[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        AddBlock(new TableBlock(rows, align, style));
        return this;
    }

    internal static TableBlock CreateTableBlockWithLinks(System.Collections.Generic.IEnumerable<string[]> rows, System.Collections.Generic.Dictionary<(int Row, int Col), string> links, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        var tb = new TableBlock(rows, align, style);
        if (links != null) {
            foreach (var kv in links) {
                if (kv.Key.Row < 0 || kv.Key.Col < 0) {
                    throw new System.ArgumentOutOfRangeException(nameof(links), "Table link row and column indexes must be non-negative.");
                }

                if (kv.Key.Row >= tb.Rows.Count) {
                    throw new System.ArgumentOutOfRangeException(nameof(links), "Table link row index must refer to an existing table row.");
                }

                if (kv.Key.Col >= tb.Rows[kv.Key.Row].Length) {
                    throw new System.ArgumentOutOfRangeException(nameof(links), "Table link column index must refer to an existing cell in the target row.");
                }

                Guard.AbsoluteUri(kv.Value, nameof(links));
                tb.AddLink(kv.Key, kv.Value);
            }
        }

        return tb;
    }

    /// <summary>
    /// Adds a table and attaches link URIs to specific cells.
    /// </summary>
    public PdfDoc TableWithLinks(System.Collections.Generic.IEnumerable<string[]> rows, System.Collections.Generic.Dictionary<(int Row, int Col), string> links, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        AddBlock(CreateTableBlockWithLinks(rows, links, align, style));
        return this;
    }

    internal static ShapeBlock CreateShapeBlock(OfficeShape shape, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        Guard.NotNull(shape, nameof(shape));
        Guard.Positive(shape.Width, nameof(shape.Width));
        Guard.Positive(shape.Height, nameof(shape.Height));
        Guard.NonNegative(shape.StrokeWidth, nameof(shape.StrokeWidth));
        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
        ValidateOpacity(shape.FillOpacity, nameof(shape.FillOpacity));
        ValidateOpacity(shape.StrokeOpacity, nameof(shape.StrokeOpacity));
        ValidateShapeClipPath(shape);
        if (shape.Kind != OfficeShapeKind.Line && shape.Kind != OfficeShapeKind.Rectangle && shape.Kind != OfficeShapeKind.RoundedRectangle && shape.Kind != OfficeShapeKind.Ellipse && shape.Kind != OfficeShapeKind.Polygon && shape.Kind != OfficeShapeKind.Path) {
            throw new System.NotSupportedException($"OfficeIMO.Pdf currently supports {nameof(OfficeShapeKind.Line)}, {nameof(OfficeShapeKind.Rectangle)}, {nameof(OfficeShapeKind.RoundedRectangle)}, {nameof(OfficeShapeKind.Ellipse)}, {nameof(OfficeShapeKind.Polygon)}, and {nameof(OfficeShapeKind.Path)} shapes only.");
        }

        if (shape.Kind == OfficeShapeKind.Line) {
            if (shape.Points.Count != 2 || shape.Points[0] == shape.Points[1]) {
                throw new System.ArgumentException("Line shapes require exactly two different points.", nameof(shape));
            }

            for (int i = 0; i < shape.Points.Count; i++) {
                ValidatePointInsideShape(shape.Points[i], shape);
            }
        }

        if (shape.Kind == OfficeShapeKind.RoundedRectangle) {
            Guard.NonNegative(shape.CornerRadius, nameof(shape.CornerRadius));
            if (shape.CornerRadius > System.Math.Min(shape.Width, shape.Height) / 2D) {
                throw new System.ArgumentOutOfRangeException(nameof(shape), "Rounded rectangle corner radius cannot exceed half of the shape width or height.");
            }
        }

        if (shape.Kind == OfficeShapeKind.Polygon) {
            if (shape.Points.Count < 3) {
                throw new System.ArgumentException("Polygon shapes require at least three points.", nameof(shape));
            }

            for (int i = 0; i < shape.Points.Count; i++) {
                var point = shape.Points[i];
                Guard.NonNegative(point.X, nameof(shape.Points));
                Guard.NonNegative(point.Y, nameof(shape.Points));
                if (point.X > shape.Width || point.Y > shape.Height) {
                    throw new System.ArgumentOutOfRangeException(nameof(shape), "Polygon points must fit inside the shape width and height.");
                }
            }
        }

        if (shape.Kind == OfficeShapeKind.Path) {
            if (shape.PathCommands.Count == 0 || shape.PathCommands[0].Kind != OfficePathCommandKind.MoveTo) {
                throw new System.ArgumentException("Path shapes require commands starting with MoveTo.", nameof(shape));
            }

            bool hasDraw = false;
            for (int i = 0; i < shape.PathCommands.Count; i++) {
                var command = shape.PathCommands[i];
                switch (command.Kind) {
                    case OfficePathCommandKind.MoveTo:
                        ValidatePointInsideShape(command.Point, shape);
                        break;
                    case OfficePathCommandKind.LineTo:
                        ValidatePointInsideShape(command.Point, shape);
                        hasDraw = true;
                        break;
                    case OfficePathCommandKind.CubicBezierTo:
                        ValidatePointInsideShape(command.ControlPoint1, shape);
                        ValidatePointInsideShape(command.ControlPoint2, shape);
                        ValidatePointInsideShape(command.Point, shape);
                        hasDraw = true;
                        break;
                    case OfficePathCommandKind.Close:
                        break;
                    default:
                        throw new System.ArgumentOutOfRangeException(nameof(shape), "Unsupported path command kind.");
                }
            }

            if (!hasDraw) {
                throw new System.ArgumentException("Path shapes require at least one drawing command.", nameof(shape));
            }
        }

        PdfDrawingStyle? drawingStyle = CreateDrawingStyle(align, spacingBefore, spacingAfter, style, "Shape");
        if (drawingStyle != null) {
            ValidateDrawingStyle(drawingStyle, "Shape");
        }

        return new ShapeBlock(shape, drawingStyle, linkUri, linkContents);
    }

    internal static PdfHorizontalRuleStyle? CreateHorizontalRuleStyle(double? thickness, PdfColor? color, double? spacingBefore, double? spacingAfter, PdfHorizontalRuleStyle? style) {
        if (!thickness.HasValue && !color.HasValue && !spacingBefore.HasValue && !spacingAfter.HasValue && style == null) {
            return null;
        }

        var ruleStyle = style?.Clone() ?? new PdfHorizontalRuleStyle();
        if (thickness.HasValue) {
            ruleStyle.Thickness = thickness.Value;
        }

        if (color.HasValue) {
            ruleStyle.Color = color.Value;
        }

        if (spacingBefore.HasValue) {
            ruleStyle.SpacingBefore = spacingBefore.Value;
        }

        if (spacingAfter.HasValue) {
            ruleStyle.SpacingAfter = spacingAfter.Value;
        }

        return ruleStyle;
    }

    internal static PdfImageStyle? CreateImageStyle(PdfAlign? align, OfficeClipPath? clipPath, OfficeImageFit? fit, double? spacingBefore, double? spacingAfter, PdfImageStyle? style) {
        if (!align.HasValue && clipPath == null && !fit.HasValue && !spacingBefore.HasValue && !spacingAfter.HasValue && style == null) {
            return null;
        }

        var imageStyle = style?.Clone() ?? new PdfImageStyle();
        if (align.HasValue) {
            imageStyle.Align = align.Value;
        }

        if (clipPath != null) {
            imageStyle.ClipPath = clipPath;
        }

        if (fit.HasValue) {
            ValidateImageFit(fit.Value, nameof(fit));
            imageStyle.Fit = fit.Value;
        }

        if (spacingBefore.HasValue) {
            imageStyle.SpacingBefore = spacingBefore.Value;
        }

        if (spacingAfter.HasValue) {
            imageStyle.SpacingAfter = spacingAfter.Value;
        }

        return imageStyle;
    }

    internal static PdfDrawingStyle? CreateDrawingStyle(PdfAlign? align, double? spacingBefore, double? spacingAfter, PdfDrawingStyle? style, string objectName = "Drawing") {
        if (!align.HasValue && !spacingBefore.HasValue && !spacingAfter.HasValue && style == null) {
            return null;
        }

        var drawingStyle = style?.Clone() ?? new PdfDrawingStyle();
        if (align.HasValue) {
            Guard.LeftCenterRightAlign(align.Value, nameof(align), objectName);
            drawingStyle.Align = align.Value;
        }

        if (spacingBefore.HasValue) {
            if (spacingBefore.Value < 0 || double.IsNaN(spacingBefore.Value) || double.IsInfinity(spacingBefore.Value)) {
                throw new System.ArgumentException(objectName + " spacing before must be a non-negative finite value.", nameof(spacingBefore));
            }

            drawingStyle.SpacingBefore = spacingBefore.Value;
        }

        if (spacingAfter.HasValue) {
            if (spacingAfter.Value < 0 || double.IsNaN(spacingAfter.Value) || double.IsInfinity(spacingAfter.Value)) {
                throw new System.ArgumentException(objectName + " spacing after must be a non-negative finite value.", nameof(spacingAfter));
            }

            drawingStyle.SpacingAfter = spacingAfter.Value;
        }

        return drawingStyle;
    }

    internal static DrawingBlock CreateDrawingBlock(OfficeDrawing drawing, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        Guard.NotNull(drawing, nameof(drawing));
        Guard.Positive(drawing.Width, nameof(drawing.Width));
        Guard.Positive(drawing.Height, nameof(drawing.Height));
        Guard.OptionalAbsoluteUri(linkUri, nameof(linkUri));
        if (drawing.Shapes.Count == 0) {
            throw new System.ArgumentException("Drawing scenes require at least one shape.", nameof(drawing));
        }

        for (int i = 0; i < drawing.Shapes.Count; i++) {
            var item = drawing.Shapes[i];
            Guard.NotNull(item, nameof(drawing.Shapes));
            Guard.NonNegative(item.X, nameof(drawing.Shapes));
            Guard.NonNegative(item.Y, nameof(drawing.Shapes));
            CreateShapeBlock(item.Shape, PdfAlign.Left, 0, 0);

            if (item.X + item.Shape.Width > drawing.Width || item.Y + item.Shape.Height > drawing.Height) {
                throw new System.ArgumentOutOfRangeException(nameof(drawing), "Drawing scene shapes must fit inside the drawing width and height.");
            }
        }

        PdfDrawingStyle? drawingStyle = CreateDrawingStyle(align, spacingBefore, spacingAfter, style, "Drawing");
        if (drawingStyle != null) {
            ValidateDrawingStyle(drawingStyle, "Drawing");
        }

        return new DrawingBlock(drawing, drawingStyle, linkUri, linkContents);
    }

    internal static void ValidateClipPathForImage(OfficeClipPath? clipPath, double width, double height) {
        ValidateClipPathInside(clipPath, width, height, nameof(clipPath), "Clip paths must fit inside the image width and height.");
    }

    internal static void ValidateImageStyleForBox(PdfImageStyle style, double width, double height, string clipPathParamName) {
        Guard.NotNull(style, nameof(style));
        Guard.LeftCenterRightAlign(style.Align, nameof(style.Align), "Image");
        ValidateImageFit(style.Fit, nameof(style.Fit));
        ValidateClipPathInside(style.ClipPath, width, height, clipPathParamName, "Clip paths must fit inside the image width and height.");
        if (style.SpacingBefore < 0 || double.IsNaN(style.SpacingBefore) || double.IsInfinity(style.SpacingBefore)) {
            throw new System.ArgumentException("Image spacing before must be a non-negative finite value.", nameof(style));
        }

        if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
            throw new System.ArgumentException("Image spacing after must be a non-negative finite value.", nameof(style));
        }
    }

    internal static void ValidateDrawingStyle(PdfDrawingStyle style, string objectName) {
        Guard.NotNull(style, nameof(style));
        Guard.LeftCenterRightAlign(style.Align, nameof(style.Align), objectName);
        if (style.SpacingBefore < 0 || double.IsNaN(style.SpacingBefore) || double.IsInfinity(style.SpacingBefore)) {
            throw new System.ArgumentException(objectName + " spacing before must be a non-negative finite value.", nameof(style));
        }

        if (style.SpacingAfter < 0 || double.IsNaN(style.SpacingAfter) || double.IsInfinity(style.SpacingAfter)) {
            throw new System.ArgumentException(objectName + " spacing after must be a non-negative finite value.", nameof(style));
        }
    }

    internal static void ValidateImageFit(OfficeImageFit fit, string paramName) {
        if (fit != OfficeImageFit.Stretch && fit != OfficeImageFit.Contain && fit != OfficeImageFit.Cover) {
            throw new System.ArgumentOutOfRangeException(paramName, "Unsupported image fit mode.");
        }
    }

    internal static void ValidateImageFitDimensions(OfficeImageInfo imageInfo, OfficeImageFit fit, string paramName) {
        if (fit == OfficeImageFit.Stretch) {
            return;
        }

        if (imageInfo.Width <= 0 || imageInfo.Height <= 0) {
            throw new System.ArgumentException("Contain and cover image fitting require image dimensions.", paramName);
        }
    }

    private static void ValidatePointInsideShape(OfficePoint point, OfficeShape shape) {
        Guard.NonNegative(point.X, nameof(shape.Points));
        Guard.NonNegative(point.Y, nameof(shape.Points));
        if (point.X > shape.Width || point.Y > shape.Height) {
            throw new System.ArgumentOutOfRangeException(nameof(shape), "Shape points must fit inside the shape width and height.");
        }
    }

    private static void ValidateShapeClipPath(OfficeShape shape) {
        var clipPath = shape.ClipPath;
        ValidateClipPathInside(clipPath, shape.Width, shape.Height, nameof(shape), "Clip paths must fit inside the shape width and height.");
    }

    private static void ValidateClipPathInside(OfficeClipPath? clipPath, double width, double height, string paramName, string fitMessage) {
        if (clipPath == null) {
            return;
        }

        Guard.Positive(clipPath.Width, paramName);
        Guard.Positive(clipPath.Height, paramName);
        if (clipPath.Width > width || clipPath.Height > height) {
            throw new System.ArgumentOutOfRangeException(paramName, fitMessage);
        }

        if (clipPath.Kind == OfficeClipPathKind.RoundedRectangle) {
            Guard.NonNegative(clipPath.CornerRadius, paramName);
            if (clipPath.CornerRadius > System.Math.Min(clipPath.Width, clipPath.Height) / 2D) {
                throw new System.ArgumentOutOfRangeException(paramName, "Clip path corner radius cannot exceed half of the clip path width or height.");
            }
        } else if (clipPath.Kind == OfficeClipPathKind.Path) {
            if (clipPath.Commands.Count == 0 || clipPath.Commands[0].Kind != OfficePathCommandKind.MoveTo) {
                throw new System.ArgumentException("Clip paths require commands starting with MoveTo.", paramName);
            }

            bool hasDraw = false;
            for (int i = 0; i < clipPath.Commands.Count; i++) {
                var command = clipPath.Commands[i];
                switch (command.Kind) {
                    case OfficePathCommandKind.MoveTo:
                        ValidatePointInsideClip(command.Point, clipPath);
                        break;
                    case OfficePathCommandKind.LineTo:
                        ValidatePointInsideClip(command.Point, clipPath);
                        hasDraw = true;
                        break;
                    case OfficePathCommandKind.CubicBezierTo:
                        ValidatePointInsideClip(command.ControlPoint1, clipPath);
                        ValidatePointInsideClip(command.ControlPoint2, clipPath);
                        ValidatePointInsideClip(command.Point, clipPath);
                        hasDraw = true;
                        break;
                    case OfficePathCommandKind.Close:
                        break;
                    default:
                        throw new System.ArgumentOutOfRangeException(paramName, "Unsupported clip path command kind.");
                }
            }

            if (!hasDraw) {
                throw new System.ArgumentException("Clip paths require at least one drawing command.", paramName);
            }
        } else if (clipPath.Kind != OfficeClipPathKind.Rectangle) {
            throw new System.ArgumentOutOfRangeException(paramName, "Unsupported clip path kind.");
        }
    }

    private static void ValidatePointInsideClip(OfficePoint point, OfficeClipPath clipPath) {
        Guard.NonNegative(point.X, nameof(clipPath.Commands));
        Guard.NonNegative(point.Y, nameof(clipPath.Commands));
        if (point.X > clipPath.Width || point.Y > clipPath.Height) {
            throw new System.ArgumentOutOfRangeException(nameof(clipPath), "Clip path commands must fit inside the clip path width and height.");
        }
    }

    private static void ValidateOpacity(double? opacity, string paramName) {
        if (!opacity.HasValue) {
            return;
        }

        double value = opacity.Value;
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value > 1D) {
            throw new System.ArgumentOutOfRangeException(paramName, "Opacity must be a finite number between 0 and 1.");
        }
    }
}

public sealed partial class PdfDoc {
    internal static OfficeImageInfo ValidateImageBytes(byte[] data) {
        if (OfficeImageReader.TryIdentify(data, null, out var info)) {
            if (info.Format == OfficeImageFormat.Jpeg) {
                return info;
            }

            if (info.Format == OfficeImageFormat.Png) {
                string? unsupportedReason;
                if (PdfWriter.TryGetPngImageData(data, out _, out unsupportedReason)) {
                    return info;
                }

                throw new NotSupportedException(
                    "PdfDoc.Image currently supports JPEG and non-interlaced 8-bit grayscale/grayscale-alpha/RGB/RGBA PNG image bytes only. " +
                    unsupportedReason);
            } else {
                throw new NotSupportedException(
                    $"PdfDoc.Image currently supports JPEG and non-interlaced 8-bit grayscale/grayscale-alpha/RGB/RGBA PNG image bytes only. Detected {info.Format} ({info.MimeType}).");
            }
        }

        if (!LooksLikeJpeg(data)) {
            System.Diagnostics.Trace.TraceWarning("PdfDoc.Image: Provided bytes do not appear to be JPEG encoded.");
        }

        return new OfficeImageInfo(OfficeImageFormat.Unknown, 0, 0);
    }

    private static bool LooksLikeJpeg(byte[] data) {
        if (data.Length < 4)
            return false;

        return data[0] == 0xFF && data[1] == 0xD8 && data[data.Length - 2] == 0xFF && data[data.Length - 1] == 0xD9;
    }
}

