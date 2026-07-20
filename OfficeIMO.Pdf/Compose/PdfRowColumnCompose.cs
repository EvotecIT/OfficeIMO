using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Column content builder used within <see cref="PdfRowCompose"/>.</summary>
public class PdfRowColumnCompose {
    private readonly RowColumn _col;
    internal PdfRowColumnCompose(RowColumn col) { _col = col; }
    /// <summary>Adds one or more flow blocks to this row column.</summary>
    public PdfRowColumnCompose Item(System.Action<PdfRowColumnCompose> build) {
        Guard.NotNull(build, nameof(build));
        build(this);
        return this;
    }
    /// <summary>Adds invisible vertical space inside the column flow.</summary>
    public PdfRowColumnCompose Spacer(double height) { _col.AddBlock(new SpacerBlock(height)); return this; }
    /// <summary>Adds an H1 heading in the column.</summary>
    public PdfRowColumnCompose H1(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null, string? linkDestinationName = null) { _col.AddBlock(new HeadingBlock(1, text, PdfAlign.Left, null, linkUri, style, linkContents, linkDestinationName)); return this; }
    /// <summary>Adds an H1 heading in the column with explicit alignment and color.</summary>
    public PdfRowColumnCompose H1(string text, PdfAlign align, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) { _col.AddBlock(new HeadingBlock(1, text, align, color, linkUri, style, linkContents, linkDestinationName)); return this; }
    /// <summary>Adds an H2 heading in the column.</summary>
    public PdfRowColumnCompose H2(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null, string? linkDestinationName = null) { _col.AddBlock(new HeadingBlock(2, text, PdfAlign.Left, null, linkUri, style, linkContents, linkDestinationName)); return this; }
    /// <summary>Adds an H2 heading in the column with explicit alignment and color.</summary>
    public PdfRowColumnCompose H2(string text, PdfAlign align, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) { _col.AddBlock(new HeadingBlock(2, text, align, color, linkUri, style, linkContents, linkDestinationName)); return this; }
    /// <summary>Adds an H3 heading in the column.</summary>
    public PdfRowColumnCompose H3(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null, string? linkDestinationName = null) { _col.AddBlock(new HeadingBlock(3, text, PdfAlign.Left, null, linkUri, style, linkContents, linkDestinationName)); return this; }
    /// <summary>Adds an H3 heading in the column with explicit alignment and color.</summary>
    public PdfRowColumnCompose H3(string text, PdfAlign align, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) { _col.AddBlock(new HeadingBlock(3, text, align, color, linkUri, style, linkContents, linkDestinationName)); return this; }
    /// <summary>Adds a page break in the column flow.</summary>
    public PdfRowColumnCompose PageBreak() { _col.AddBlock(new PageBreakBlock()); return this; }
    /// <summary>Adds a paragraph built from styled runs to the column.</summary>
    /// <param name="build">Paragraph content builder.</param>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    /// <param name="style">Optional paragraph layout style.</param>
    public PdfRowColumnCompose Paragraph(System.Action<PdfParagraphBuilder> build, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null, PdfParagraphStyle? style = null) {
        Guard.NotNull(build, nameof(build));
        var b = new PdfParagraphBuilder(align, defaultColor);
        build(b);
        _col.AddBlock(new RichParagraphBlock(b.Build().Runs, align, defaultColor, style));
        return this;
    }
    /// <summary>Adds a simple bullet list in the column.</summary>
    public PdfRowColumnCompose Bullets(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, PdfListStyle? style = null) {
        _col.AddBlock(new BulletListBlock(items, align, color, style));
        return this;
    }
    /// <summary>Adds a bullet list whose items can contain rich inline text runs in the column.</summary>
    public PdfRowColumnCompose RichBullets(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, PdfListStyle? style = null) {
        _col.AddBlock(new BulletListBlock(items, align, color, style));
        return this;
    }
    /// <summary>Adds a simple numbered list in the column.</summary>
    public PdfRowColumnCompose Numbered(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, int startNumber = 1, PdfListStyle? style = null) {
        _col.AddBlock(new NumberedListBlock(items, align, color, startNumber, style));
        return this;
    }
    /// <summary>Adds a numbered list whose items can contain rich inline text runs in the column.</summary>
    public PdfRowColumnCompose RichNumbered(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, int startNumber = 1, PdfListStyle? style = null) {
        _col.AddBlock(new NumberedListBlock(items, align, color, startNumber, style));
        return this;
    }
    /// <summary>Adds a paragraph inside a styled panel in the column.</summary>
    public PdfRowColumnCompose PanelParagraph(System.Action<PdfParagraphBuilder> build, PanelStyle? style = null, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) {
        Guard.NotNull(build, nameof(build));
        var builder = new PdfParagraphBuilder(align, defaultColor);
        build(builder);
        _col.AddBlock(new PanelParagraphBlock(builder.Build().Runs, align, defaultColor, style));
        return this;
    }
    /// <summary>Adds a styled panel from common column flow blocks such as paragraphs, headings, lists, simple tables, rules, and nested panel paragraphs.</summary>
    public PdfRowColumnCompose Panel(System.Action<PdfRowColumnCompose> build, PanelStyle? style = null, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) {
        Guard.NotNull(build, nameof(build));
        Guard.ParagraphAlign(align, nameof(align), "Panel");
        var panelColumn = new RowColumn(100);
        build(new PdfRowColumnCompose(panelColumn));
        _col.AddBlock(PdfDocument.CreatePanelParagraphBlock(panelColumn.Blocks, style, align, defaultColor));
        return this;
    }
    /// <summary>Adds a simple table in the column.</summary>
    public PdfRowColumnCompose Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        _col.AddBlock(new TableBlock(rows, align, style));
        return this;
    }
    /// <summary>Adds a table with explicit cells, including optional column spans.</summary>
    public PdfRowColumnCompose Table(System.Collections.Generic.IEnumerable<PdfTableCell[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        _col.AddBlock(new TableBlock(rows, align, style));
        return this;
    }
    /// <summary>Adds a simple table in the column and attaches link URIs to specific cells.</summary>
    public PdfRowColumnCompose TableWithLinks(System.Collections.Generic.IEnumerable<string[]> rows, System.Collections.Generic.Dictionary<(int Row, int Col), string> links, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) {
        _col.AddBlock(PdfDocument.CreateTableBlockWithLinks(rows, links, align, style));
        return this;
    }
    /// <summary>Adds a horizontal rule in the column.</summary>
    public PdfRowColumnCompose HR(double? thickness = null, PdfColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfHorizontalRuleStyle? style = null) {
        _col.AddBlock(new HorizontalRuleBlock(PdfDocument.CreateHorizontalRuleStyle(thickness, color, spacingBefore, spacingAfter, style)));
        return this;
    }
    /// <summary>Adds a named bookmark at the current column flow position.</summary>
    public PdfRowColumnCompose Bookmark(string name) { _col.AddBlock(new BookmarkBlock(name)); return this; }
    /// <summary>Adds a simple AcroForm text field in the column.</summary>
    public PdfRowColumnCompose TextField(string name, double width = 180, double height = 22, string value = "", PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) {
        _col.AddBlock(new TextFieldBlock(name, width, height, value, align, fontSize, spacingBefore, spacingAfter, style));
        return this;
    }
    /// <summary>Adds a simple AcroForm check box in the column.</summary>
    public PdfRowColumnCompose CheckBox(string name, bool isChecked = false, double size = 14, PdfAlign align = PdfAlign.Left, double spacingBefore = 0, double spacingAfter = 6, string checkedValueName = "Yes", PdfFormFieldStyle? style = null) {
        _col.AddBlock(new CheckBoxBlock(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style));
        return this;
    }
    /// <summary>Adds a simple AcroForm choice field in the column.</summary>
    public PdfRowColumnCompose ChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, double width = 180, double height = 22, PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, bool isComboBox = true, PdfFormFieldStyle? style = null) {
        _col.AddBlock(new ChoiceFieldBlock(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style));
        return this;
    }
    /// <summary>Adds a simple AcroForm multi-select choice field in the column.</summary>
    public PdfRowColumnCompose MultiSelectChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, System.Collections.Generic.IEnumerable<string>? values = null, double width = 180, double height = 72, PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) {
        _col.AddBlock(new ChoiceFieldBlock(name, options, values, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox: false, allowsMultipleSelection: true, style));
        return this;
    }
    /// <summary>Adds a simple AcroForm radio button group in the column.</summary>
    public PdfRowColumnCompose RadioButtonGroup(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, double size = 14, double gap = 6, PdfAlign align = PdfAlign.Left, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) {
        _col.AddBlock(new RadioButtonGroupBlock(name, options, value, size, gap, align, spacingBefore, spacingAfter, style));
        return this;
    }
    /// <summary>Adds a PDF text annotation in the column.</summary>
    public PdfRowColumnCompose TextAnnotation(string contents, double width = 18D, double height = 18D, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfTextAnnotationIcon icon = PdfTextAnnotationIcon.Comment, PdfColor? color = null, bool open = false) {
        _col.AddBlock(PdfDocument.CreateTextAnnotationBlock(contents, width, height, align, spacingBefore, spacingAfter, icon, color, open));
        return this;
    }
    /// <summary>Adds a PDF free-text annotation in the column.</summary>
    public PdfRowColumnCompose FreeTextAnnotation(string contents, double width, double height, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, double fontSize = 10D, PdfColor? textColor = null, PdfColor? borderColor = null, double borderWidth = 1D, PdfColor? fillColor = null, PdfAlign textAlign = PdfAlign.Left, double padding = 3D, double? lineHeight = null) {
        _col.AddBlock(PdfDocument.CreateFreeTextAnnotationBlock(contents, width, height, align, spacingBefore, spacingAfter, fontSize, textColor, borderColor, borderWidth, fillColor, textAlign, padding, lineHeight));
        return this;
    }
    /// <summary>Adds a PDF highlight annotation rectangle in the column.</summary>
    public PdfRowColumnCompose HighlightAnnotation(string contents, double width, double height, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfColor? color = null) {
        _col.AddBlock(PdfDocument.CreateHighlightAnnotationBlock(contents, width, height, align, spacingBefore, spacingAfter, color));
        return this;
    }
    /// <summary>Adds a shared OfficeIMO.Drawing shape in the column.</summary>
    public PdfRowColumnCompose Shape(OfficeShape shape, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _col.AddBlock(PdfDocument.CreateShapeBlock(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents)); return this; }
    /// <summary>Adds a shared OfficeIMO.Drawing scene in the column.</summary>
    public PdfRowColumnCompose Drawing(OfficeDrawing drawing, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _col.AddBlock(PdfDocument.CreateDrawingBlock(drawing, align, spacingBefore, spacingAfter, style, linkUri, linkContents)); return this; }
    /// <summary>Adds a line vector shape in the column.</summary>
    public PdfRowColumnCompose Line(double x1, double y1, double x2, double y2, PdfColor? strokeColor = null, double strokeWidth = 1, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Line(x1, y1, x2, y2);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        _col.AddBlock(PdfDocument.CreateShapeBlock(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents));
        return this;
    }
    /// <summary>Adds a rectangle vector shape in the column.</summary>
    public PdfRowColumnCompose Rectangle(double width, double height, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Rectangle(width, height);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        _col.AddBlock(PdfDocument.CreateShapeBlock(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents));
        return this;
    }
    /// <summary>Adds a rounded rectangle vector shape in the column.</summary>
    public PdfRowColumnCompose RoundedRectangle(double width, double height, double cornerRadius, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.RoundedRectangle(width, height, cornerRadius);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        _col.AddBlock(PdfDocument.CreateShapeBlock(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents));
        return this;
    }
    /// <summary>Adds an ellipse vector shape in the column.</summary>
    public PdfRowColumnCompose Ellipse(double width, double height, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Ellipse(width, height);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        _col.AddBlock(PdfDocument.CreateShapeBlock(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents));
        return this;
    }
    /// <summary>Adds a polygon vector shape in the column.</summary>
    public PdfRowColumnCompose Polygon(System.Collections.Generic.IEnumerable<OfficePoint> points, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Polygon(points);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        _col.AddBlock(PdfDocument.CreateShapeBlock(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents));
        return this;
    }
    /// <summary>Adds a freeform path vector shape in the column.</summary>
    public PdfRowColumnCompose Path(System.Collections.Generic.IEnumerable<OfficePathCommand> commands, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Path(commands);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        _col.AddBlock(PdfDocument.CreateShapeBlock(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents));
        return this;
    }
    /// <summary>Adds a raster image supported by OfficeIMO.Drawing in the column.</summary>
    public PdfRowColumnCompose Image(byte[] jpegBytes, double width, double height, PdfAlign? align = null, OfficeClipPath? clipPath = null, OfficeImageFit? fit = null, double? spacingBefore = null, double? spacingAfter = null, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null) =>
        Image(jpegBytes, width, height, align, clipPath, fit, spacingBefore, spacingAfter, style, linkUri, linkContents, alternativeText: null);

    /// <summary>Adds a supported meaningful image in the column with alternate text.</summary>
    public PdfRowColumnCompose Image(byte[] jpegBytes, double width, double height, string? alternativeText) =>
        Image(jpegBytes, width, height, align: null, clipPath: null, fit: null, spacingBefore: null, spacingAfter: null, style: null, linkUri: null, linkContents: null, alternativeText: alternativeText);

    /// <summary>Adds a raster image supported by OfficeIMO.Drawing in the column.</summary>
    public PdfRowColumnCompose Image(byte[] jpegBytes, double width, double height, PdfAlign? align, OfficeClipPath? clipPath, OfficeImageFit? fit, double? spacingBefore, double? spacingAfter, PdfImageStyle? style, string? linkUri, string? linkContents, string? alternativeText) {
        Guard.NotNullOrEmpty(jpegBytes, nameof(jpegBytes));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.OptionalUriAction(linkUri, nameof(linkUri));
        PdfImageStyle? imageStyle = PdfDocument.CreateImageStyle(align, clipPath, fit, spacingBefore, spacingAfter, style, alternativeText);
        if (imageStyle != null) {
            PdfDocument.ValidateImageStyleForBox(imageStyle, width, height, nameof(clipPath));
        }

        PdfDocument.PreparedImage prepared = PdfDocument.PrepareImageBytes(jpegBytes);
        if (imageStyle != null) {
            PdfDocument.ValidateImageFitDimensions(prepared.Info, imageStyle.Fit, nameof(fit));
        }

        _col.AddBlock(new ImageBlock(prepared.Data, width, height, prepared.Info, imageStyle, linkUri, linkContents, useDataSnapshot: true));
        return this;
    }
}
