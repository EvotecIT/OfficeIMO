using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Builder for individual flow items (headings, paragraphs, tables, images).</summary>
public class PdfItemCompose {
    private readonly PdfDoc _doc;
    internal PdfItemCompose(PdfDoc doc) { _doc = doc; }
    /// <summary>Starts a new page.</summary>
    public PdfItemCompose PageBreak() { _doc.PageBreak(); return this; }
    /// <summary>Adds invisible vertical space to the current flow.</summary>
    public PdfItemCompose Spacer(double height) { _doc.Spacer(height); return this; }
    /// <summary>Adds an H1 heading.</summary>
    public PdfItemCompose H1(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.H1(text, style: style, linkUri: linkUri, linkContents: linkContents); return this; }
    /// <summary>Adds an H1 heading with explicit alignment and color.</summary>
    public PdfItemCompose H1(string text, PdfAlign align, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null) { _doc.H1(text, align, color, linkUri, style, linkContents); return this; }
    /// <summary>Adds an H2 heading.</summary>
    public PdfItemCompose H2(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.H2(text, style: style, linkUri: linkUri, linkContents: linkContents); return this; }
    /// <summary>Adds an H2 heading with explicit alignment and color.</summary>
    public PdfItemCompose H2(string text, PdfAlign align, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null) { _doc.H2(text, align, color, linkUri, style, linkContents); return this; }
    /// <summary>Adds an H3 heading.</summary>
    public PdfItemCompose H3(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.H3(text, style: style, linkUri: linkUri, linkContents: linkContents); return this; }
    /// <summary>Adds an H3 heading with explicit alignment and color.</summary>
    public PdfItemCompose H3(string text, PdfAlign align, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null) { _doc.H3(text, align, color, linkUri, style, linkContents); return this; }
    /// <summary>Adds a paragraph built from styled text runs.</summary>
    /// <param name="build">Paragraph content builder.</param>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    /// <param name="style">Optional paragraph layout style.</param>
    public PdfItemCompose Paragraph(System.Action<PdfParagraphBuilder> build, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null, PdfParagraphStyle? style = null) { _doc.Paragraph(build, align, defaultColor, style); return this; }
    /// <summary>Adds a simple bullet list.</summary>
    public PdfItemCompose Bullets(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, PdfListStyle? style = null) { _doc.Bullets(items, align, color, style); return this; }
    /// <summary>Adds a bullet list whose items can contain rich inline text runs.</summary>
    public PdfItemCompose RichBullets(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, PdfListStyle? style = null) { _doc.RichBullets(items, align, color, style); return this; }
    /// <summary>Adds a simple numbered list.</summary>
    public PdfItemCompose Numbered(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, int startNumber = 1, PdfListStyle? style = null) { _doc.Numbered(items, align, color, startNumber, style); return this; }
    /// <summary>Adds a numbered list whose items can contain rich inline text runs.</summary>
    public PdfItemCompose RichNumbered(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, int startNumber = 1, PdfListStyle? style = null) { _doc.RichNumbered(items, align, color, startNumber, style); return this; }
    /// <summary>Adds a simple text table.</summary>
    /// <param name="rows">Sequence of row arrays.</param>
    /// <param name="align">Table alignment.</param>
    /// <param name="style">Optional table styling.</param>
    public PdfItemCompose Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.Table(rows, align, style); return this; }
    /// <summary>Adds a table with explicit cells, including optional column spans.</summary>
    /// <param name="rows">Sequence of rows made from explicit table cells.</param>
    /// <param name="align">Table alignment.</param>
    /// <param name="style">Optional table styling.</param>
    public PdfItemCompose Table(System.Collections.Generic.IEnumerable<PdfTableCell[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.Table(rows, align, style); return this; }
    /// <summary>Adds a simple text table and attaches link URIs to specific cells.</summary>
    /// <param name="rows">Sequence of row arrays.</param>
    /// <param name="links">Per-cell absolute link URIs keyed by zero-based row and column.</param>
    /// <param name="align">Table alignment.</param>
    /// <param name="style">Optional table styling.</param>
    public PdfItemCompose TableWithLinks(System.Collections.Generic.IEnumerable<string[]> rows, System.Collections.Generic.Dictionary<(int Row, int Col), string> links, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.TableWithLinks(rows, links, align, style); return this; }
    /// <summary>Builds nested elements (e.g., grouping heading + paragraph).</summary>
    /// <param name="build">Delegate composing nested elements.</param>
    public PdfItemCompose Element(System.Action<PdfElementCompose> build) { Guard.NotNull(build, nameof(build)); var el = new PdfElementCompose(_doc); build(el); return this; }
    /// <summary>Adds a horizontal rule.</summary>
    /// <param name="thickness">Line thickness (pt).</param>
    /// <param name="color">Optional color; inherited from the current default rule style when omitted.</param>
    /// <param name="spacingBefore">Top spacing (pt), inherited from the current default rule style when omitted.</param>
    /// <param name="spacingAfter">Bottom spacing (pt), inherited from the current default rule style when omitted.</param>
    /// <param name="style">Optional reusable rule style.</param>
    public PdfItemCompose HR(double? thickness = null, PdfColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfHorizontalRuleStyle? style = null) { _doc.HR(thickness, color, spacingBefore, spacingAfter, style); return this; }
    /// <summary>Adds a named bookmark at the current flow position.</summary>
    public PdfItemCompose Bookmark(string name) { _doc.Bookmark(name); return this; }
    /// <summary>Adds a simple AcroForm text field at the current flow position.</summary>
    public PdfItemCompose TextField(string name, double width = 180, double height = 22, string value = "", PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) { _doc.TextField(name, width, height, value, align, fontSize, spacingBefore, spacingAfter, style); return this; }
    /// <summary>Adds a simple AcroForm check box at the current flow position.</summary>
    public PdfItemCompose CheckBox(string name, bool isChecked = false, double size = 14, PdfAlign align = PdfAlign.Left, double spacingBefore = 0, double spacingAfter = 6, string checkedValueName = "Yes", PdfFormFieldStyle? style = null) { _doc.CheckBox(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style); return this; }
    /// <summary>Adds a simple AcroForm choice field at the current flow position.</summary>
    public PdfItemCompose ChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, double width = 180, double height = 22, PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, bool isComboBox = true, PdfFormFieldStyle? style = null) { _doc.ChoiceField(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style); return this; }
    /// <summary>Adds a simple AcroForm multi-select choice field at the current flow position.</summary>
    public PdfItemCompose MultiSelectChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, System.Collections.Generic.IEnumerable<string>? values = null, double width = 180, double height = 72, PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) { _doc.MultiSelectChoiceField(name, options, values, width, height, align, fontSize, spacingBefore, spacingAfter, style); return this; }
    /// <summary>Adds a simple AcroForm radio button group at the current flow position.</summary>
    public PdfItemCompose RadioButtonGroup(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, double size = 14, double gap = 6, PdfAlign align = PdfAlign.Left, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) { _doc.RadioButtonGroup(name, options, value, size, gap, align, spacingBefore, spacingAfter, style); return this; }
    /// <summary>Adds a shared OfficeIMO.Drawing shape.</summary>
    public PdfItemCompose Shape(OfficeShape shape, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents); return this; }
    /// <summary>Adds a shared OfficeIMO.Drawing scene.</summary>
    public PdfItemCompose Drawing(OfficeDrawing drawing, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Drawing(drawing, align, spacingBefore, spacingAfter, style, linkUri, linkContents); return this; }
    /// <summary>Adds a line vector shape.</summary>
    public PdfItemCompose Line(double x1, double y1, double x2, double y2, PdfColor? strokeColor = null, double strokeWidth = 1, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Line(x1, y1, x2, y2, strokeColor, strokeWidth, align, spacingBefore, spacingAfter, strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a rectangle vector shape.</summary>
    public PdfItemCompose Rectangle(double width, double height, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Rectangle(width, height, strokeColor, strokeWidth, fillColor, align, spacingBefore, spacingAfter, strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a rounded rectangle vector shape.</summary>
    public PdfItemCompose RoundedRectangle(double width, double height, double cornerRadius, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.RoundedRectangle(width, height, cornerRadius, strokeColor, strokeWidth, fillColor, align, spacingBefore, spacingAfter, strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds an ellipse vector shape.</summary>
    public PdfItemCompose Ellipse(double width, double height, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Ellipse(width, height, strokeColor, strokeWidth, fillColor, align, spacingBefore, spacingAfter, strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a polygon vector shape.</summary>
    public PdfItemCompose Polygon(System.Collections.Generic.IEnumerable<OfficePoint> points, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Polygon(points, strokeColor, strokeWidth, fillColor, align, spacingBefore, spacingAfter, strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a freeform path vector shape.</summary>
    public PdfItemCompose Path(System.Collections.Generic.IEnumerable<OfficePathCommand> commands, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Path(commands, strokeColor, strokeWidth, fillColor, align, spacingBefore, spacingAfter, strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a paragraph inside a styled panel.</summary>
    /// <param name="build">Paragraph content builder.</param>
    /// <param name="style">Panel style (padding, background, border, etc.).</param>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    public PdfItemCompose PanelParagraph(System.Action<PdfParagraphBuilder> build, PanelStyle? style = null, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) { _doc.PanelParagraph(build, style, align, defaultColor); return this; }
    /// <summary>Adds a styled panel from common flow blocks such as paragraphs, headings, lists, simple tables, rules, and nested panel paragraphs.</summary>
    /// <param name="build">Panel content builder.</param>
    /// <param name="style">Panel style (padding, background, border, etc.).</param>
    /// <param name="align">Panel text alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    public PdfItemCompose Panel(System.Action<PdfItemCompose> build, PanelStyle? style = null, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) { _doc.Panel(build, style, align, defaultColor); return this; }
    /// <summary>Adds an image from supported image bytes. JPEG and simple non-interlaced 8-bit PNG images, including grayscale-alpha/RGBA soft masks, are currently supported.</summary>
    /// <param name="jpegBytes">Supported image bytes.</param>
    /// <param name="width">Target width in points.</param>
    /// <param name="height">Target height in points.</param>
    /// <param name="align">Image alignment inside content width.</param>
    /// <param name="clipPath">Optional local clipping path applied before drawing the image.</param>
    /// <param name="fit">Image fitting mode inside the target box.</param>
    /// <param name="spacingBefore">Top spacing (pt), inherited from the current default image style when omitted.</param>
    /// <param name="spacingAfter">Bottom spacing (pt), inherited from the current default image style when omitted.</param>
    /// <param name="style">Optional reusable image placement style.</param>
    /// <param name="linkUri">Optional absolute URI for an image link annotation.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfItemCompose Image(byte[] jpegBytes, double width, double height, PdfAlign? align = null, OfficeClipPath? clipPath = null, OfficeImageFit? fit = null, double? spacingBefore = null, double? spacingAfter = null, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null) {
        Guard.NotNullOrEmpty(jpegBytes, nameof(jpegBytes));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));

        _doc.Image(jpegBytes, width, height, align, clipPath, fit, spacingBefore, spacingAfter, style, linkUri, linkContents);
        return this;
    }
}
