using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Builder for nested elements used within item builders.</summary>
public class PdfElementCompose {
    private readonly PdfDocument _doc;
    internal PdfElementCompose(PdfDocument doc) { _doc = doc; }
    /// <summary>Adds invisible vertical space to the current flow.</summary>
    public PdfElementCompose Spacer(double height) { _doc.Spacer(height); return this; }
    /// <summary>Starts a new page from the current nested element flow.</summary>
    public PdfElementCompose PageBreak() { _doc.PageBreak(); return this; }
    /// <summary>Adds an H1 heading.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="style">Optional heading style.</param>
    /// <param name="linkUri">Optional absolute URI or catalog-base-relative URI for a heading link annotation.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfElementCompose H1(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.H1(text, style: style, linkUri: linkUri, linkContents: linkContents); return this; }
    /// <summary>Adds an H1 heading with explicit alignment and color.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="align">Heading alignment.</param>
    /// <param name="color">Optional heading color.</param>
    /// <param name="linkUri">Optional absolute URI or catalog-base-relative URI for a heading link annotation.</param>
    /// <param name="style">Optional heading style.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfElementCompose H1(string text, PdfAlign align, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null) { _doc.H1(text, align, color, linkUri, style, linkContents); return this; }
    /// <summary>Adds an H2 heading.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="style">Optional heading style.</param>
    /// <param name="linkUri">Optional absolute URI or catalog-base-relative URI for a heading link annotation.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfElementCompose H2(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.H2(text, style: style, linkUri: linkUri, linkContents: linkContents); return this; }
    /// <summary>Adds an H2 heading with explicit alignment and color.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="align">Heading alignment.</param>
    /// <param name="color">Optional heading color.</param>
    /// <param name="linkUri">Optional absolute URI or catalog-base-relative URI for a heading link annotation.</param>
    /// <param name="style">Optional heading style.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfElementCompose H2(string text, PdfAlign align, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null) { _doc.H2(text, align, color, linkUri, style, linkContents); return this; }
    /// <summary>Adds an H3 heading.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="style">Optional heading style.</param>
    /// <param name="linkUri">Optional absolute URI or catalog-base-relative URI for a heading link annotation.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfElementCompose H3(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.H3(text, style: style, linkUri: linkUri, linkContents: linkContents); return this; }
    /// <summary>Adds an H3 heading with explicit alignment and color.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="align">Heading alignment.</param>
    /// <param name="color">Optional heading color.</param>
    /// <param name="linkUri">Optional absolute URI or catalog-base-relative URI for a heading link annotation.</param>
    /// <param name="style">Optional heading style.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfElementCompose H3(string text, PdfAlign align, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null) { _doc.H3(text, align, color, linkUri, style, linkContents); return this; }
    /// <summary>Adds a paragraph built from styled text runs.</summary>
    /// <param name="build">Paragraph content builder.</param>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    /// <param name="style">Optional paragraph layout style.</param>
    public PdfElementCompose Paragraph(System.Action<PdfParagraphBuilder> build, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null, PdfParagraphStyle? style = null) { _doc.Paragraph(build, align, defaultColor, style); return this; }
    /// <summary>Adds a simple bullet list.</summary>
    public PdfElementCompose Bullets(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, PdfListStyle? style = null) { _doc.Bullets(items, align, color, style); return this; }
    /// <summary>Adds a bullet list whose items can contain rich inline text runs.</summary>
    public PdfElementCompose RichBullets(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, PdfListStyle? style = null) { _doc.RichBullets(items, align, color, style); return this; }
    /// <summary>Adds a simple numbered list.</summary>
    public PdfElementCompose Numbered(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, int startNumber = 1, PdfListStyle? style = null) { _doc.Numbered(items, align, color, startNumber, style); return this; }
    /// <summary>Adds a numbered list whose items can contain rich inline text runs.</summary>
    public PdfElementCompose RichNumbered(System.Collections.Generic.IEnumerable<PdfListItem> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, int startNumber = 1, PdfListStyle? style = null) { _doc.RichNumbered(items, align, color, startNumber, style); return this; }
    /// <summary>Adds a simple text table.</summary>
    /// <param name="rows">Sequence of row arrays.</param>
    /// <param name="align">Table alignment.</param>
    /// <param name="style">Optional table styling.</param>
    public PdfElementCompose Table(System.Collections.Generic.IEnumerable<string[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.Table(rows, align, style); return this; }
    /// <summary>Adds a table with explicit cells, including optional column spans.</summary>
    /// <param name="rows">Sequence of rows made from explicit table cells.</param>
    /// <param name="align">Table alignment.</param>
    /// <param name="style">Optional table styling.</param>
    public PdfElementCompose Table(System.Collections.Generic.IEnumerable<PdfTableCell[]> rows, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.Table(rows, align, style); return this; }
    /// <summary>Adds a simple text table and attaches link URIs to specific cells.</summary>
    /// <param name="rows">Sequence of row arrays.</param>
    /// <param name="links">Per-cell absolute or catalog-base-relative link URIs keyed by zero-based row and column.</param>
    /// <param name="align">Table alignment.</param>
    /// <param name="style">Optional table styling.</param>
    public PdfElementCompose TableWithLinks(System.Collections.Generic.IEnumerable<string[]> rows, System.Collections.Generic.Dictionary<(int Row, int Col), string> links, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.TableWithLinks(rows, links, align, style); return this; }
    /// <summary>Adds a horizontal rule at the current nested element flow position.</summary>
    /// <param name="thickness">Line thickness (pt).</param>
    /// <param name="color">Optional color; inherited from the current default rule style when omitted.</param>
    /// <param name="spacingBefore">Top spacing (pt), inherited from the current default rule style when omitted.</param>
    /// <param name="spacingAfter">Bottom spacing (pt), inherited from the current default rule style when omitted.</param>
    /// <param name="style">Optional reusable rule style.</param>
    public PdfElementCompose HR(double? thickness = null, PdfColor? color = null, double? spacingBefore = null, double? spacingAfter = null, PdfHorizontalRuleStyle? style = null) { _doc.HR(thickness, color, spacingBefore, spacingAfter, style); return this; }
    /// <summary>Adds a named bookmark at the current nested element flow position.</summary>
    public PdfElementCompose Bookmark(string name) { _doc.Bookmark(name); return this; }
    /// <summary>Adds a paragraph inside a styled panel at the current nested element flow position.</summary>
    /// <param name="build">Paragraph content builder.</param>
    /// <param name="style">Panel style (padding, background, border, etc.).</param>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    public PdfElementCompose PanelParagraph(System.Action<PdfParagraphBuilder> build, PanelStyle? style = null, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) { _doc.PanelParagraph(build, style, align, defaultColor); return this; }
    /// <summary>Adds a styled panel from common flow blocks such as paragraphs, headings, lists, simple tables, rules, and nested panel paragraphs.</summary>
    /// <param name="build">Panel content builder.</param>
    /// <param name="style">Panel style (padding, background, border, etc.).</param>
    /// <param name="align">Panel text alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    public PdfElementCompose Panel(System.Action<PdfItemCompose> build, PanelStyle? style = null, PdfAlign align = PdfAlign.Left, PdfColor? defaultColor = null) { _doc.Panel(build, style, align, defaultColor); return this; }
    /// <summary>Adds a simple AcroForm text field at the current nested element flow position.</summary>
    public PdfElementCompose TextField(string name, double width = 180, double height = 22, string value = "", PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) { _doc.TextField(name, width, height, value, align, fontSize, spacingBefore, spacingAfter, style); return this; }
    /// <summary>Adds a simple AcroForm check box at the current nested element flow position.</summary>
    public PdfElementCompose CheckBox(string name, bool isChecked = false, double size = 14, PdfAlign align = PdfAlign.Left, double spacingBefore = 0, double spacingAfter = 6, string checkedValueName = "Yes", PdfFormFieldStyle? style = null) { _doc.CheckBox(name, isChecked, size, align, spacingBefore, spacingAfter, checkedValueName, style); return this; }
    /// <summary>Adds a simple AcroForm choice field at the current nested element flow position.</summary>
    public PdfElementCompose ChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, double width = 180, double height = 22, PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, bool isComboBox = true, PdfFormFieldStyle? style = null) { _doc.ChoiceField(name, options, value, width, height, align, fontSize, spacingBefore, spacingAfter, isComboBox, style); return this; }
    /// <summary>Adds a simple AcroForm multi-select choice field at the current nested element flow position.</summary>
    public PdfElementCompose MultiSelectChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, System.Collections.Generic.IEnumerable<string>? values = null, double width = 180, double height = 72, PdfAlign align = PdfAlign.Left, double fontSize = 10, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) { _doc.MultiSelectChoiceField(name, options, values, width, height, align, fontSize, spacingBefore, spacingAfter, style); return this; }
    /// <summary>Adds a simple AcroForm radio button group at the current nested element flow position.</summary>
    public PdfElementCompose RadioButtonGroup(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, double size = 14, double gap = 6, PdfAlign align = PdfAlign.Left, double spacingBefore = 0, double spacingAfter = 6, PdfFormFieldStyle? style = null) { _doc.RadioButtonGroup(name, options, value, size, gap, align, spacingBefore, spacingAfter, style); return this; }
    /// <summary>Adds a PDF text annotation at the current nested element flow position.</summary>
    public PdfElementCompose TextAnnotation(string contents, double width = 18D, double height = 18D, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfTextAnnotationIcon icon = PdfTextAnnotationIcon.Comment, PdfColor? color = null, bool open = false) { _doc.TextAnnotation(contents, width, height, align, spacingBefore, spacingAfter, icon, color, open); return this; }
    /// <summary>Adds a PDF free-text annotation at the current nested element flow position.</summary>
    public PdfElementCompose FreeTextAnnotation(string contents, double width, double height, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, double fontSize = 10D, PdfColor? textColor = null, PdfColor? borderColor = null, double borderWidth = 1D, PdfColor? fillColor = null, PdfAlign textAlign = PdfAlign.Left, double padding = 3D, double? lineHeight = null) { _doc.FreeTextAnnotation(contents, width, height, align, spacingBefore, spacingAfter, fontSize, textColor, borderColor, borderWidth, fillColor, textAlign, padding, lineHeight); return this; }
    /// <summary>Adds a PDF highlight annotation rectangle at the current nested element flow position.</summary>
    public PdfElementCompose HighlightAnnotation(string contents, double width, double height, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfColor? color = null) { _doc.HighlightAnnotation(contents, width, height, align, spacingBefore, spacingAfter, color); return this; }
    /// <summary>Adds a shared OfficeIMO.Drawing shape.</summary>
    public PdfElementCompose Shape(OfficeShape shape, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents); return this; }
    /// <summary>Adds a shared OfficeIMO.Drawing scene.</summary>
    public PdfElementCompose Drawing(OfficeDrawing drawing, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Drawing(drawing, align, spacingBefore, spacingAfter, style, linkUri, linkContents); return this; }
    /// <summary>Adds a line vector shape.</summary>
    public PdfElementCompose Line(double x1, double y1, double x2, double y2, PdfColor? strokeColor = null, double strokeWidth = 1, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Line(x1, y1, x2, y2, strokeColor, strokeWidth, align, spacingBefore, spacingAfter, strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a rectangle vector shape.</summary>
    public PdfElementCompose Rectangle(double width, double height, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Rectangle(width, height, strokeColor, strokeWidth, fillColor, align, spacingBefore, spacingAfter, strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a rounded rectangle vector shape.</summary>
    public PdfElementCompose RoundedRectangle(double width, double height, double cornerRadius, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.RoundedRectangle(width, height, cornerRadius, strokeColor, strokeWidth, fillColor, align, spacingBefore, spacingAfter, strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds an ellipse vector shape.</summary>
    public PdfElementCompose Ellipse(double width, double height, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Ellipse(width, height, strokeColor, strokeWidth, fillColor, align, spacingBefore, spacingAfter, strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a polygon vector shape.</summary>
    public PdfElementCompose Polygon(System.Collections.Generic.IEnumerable<OfficePoint> points, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Polygon(points, strokeColor, strokeWidth, fillColor, align, spacingBefore, spacingAfter, strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a freeform path vector shape.</summary>
    public PdfElementCompose Path(System.Collections.Generic.IEnumerable<OfficePathCommand> commands, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Path(commands, strokeColor, strokeWidth, fillColor, align, spacingBefore, spacingAfter, strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a raster image supported by OfficeIMO.Drawing at the current nested element flow position.</summary>
    /// <param name="jpegBytes">Supported image bytes.</param>
    /// <param name="width">Target width in points.</param>
    /// <param name="height">Target height in points.</param>
    /// <param name="align">Image alignment inside content width.</param>
    /// <param name="clipPath">Optional local clipping path applied before drawing the image.</param>
    /// <param name="fit">Image fitting mode inside the target box.</param>
    /// <param name="spacingBefore">Top spacing (pt), inherited from the current default image style when omitted.</param>
    /// <param name="spacingAfter">Bottom spacing (pt), inherited from the current default image style when omitted.</param>
    /// <param name="style">Optional reusable image placement style.</param>
    /// <param name="linkUri">Optional absolute URI or catalog-base-relative URI for an image link annotation.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfElementCompose Image(byte[] jpegBytes, double width, double height, PdfAlign? align = null, OfficeClipPath? clipPath = null, OfficeImageFit? fit = null, double? spacingBefore = null, double? spacingAfter = null, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null) =>
        Image(jpegBytes, width, height, align, clipPath, fit, spacingBefore, spacingAfter, style, linkUri, linkContents, alternativeText: null);

    /// <summary>Adds a meaningful image from supported image bytes with alternate text.</summary>
    public PdfElementCompose Image(byte[] jpegBytes, double width, double height, string? alternativeText) =>
        Image(jpegBytes, width, height, align: null, clipPath: null, fit: null, spacingBefore: null, spacingAfter: null, style: null, linkUri: null, linkContents: null, alternativeText: alternativeText);

    /// <summary>Adds a raster image supported by OfficeIMO.Drawing.</summary>
    /// <param name="jpegBytes">Supported image bytes.</param>
    /// <param name="width">Target width in points.</param>
    /// <param name="height">Target height in points.</param>
    /// <param name="align">Image alignment inside content width.</param>
    /// <param name="clipPath">Optional local clipping path applied before drawing the image.</param>
    /// <param name="fit">Image fitting mode inside the target box.</param>
    /// <param name="spacingBefore">Top spacing (pt), inherited from the current default image style when omitted.</param>
    /// <param name="spacingAfter">Bottom spacing (pt), inherited from the current default image style when omitted.</param>
    /// <param name="style">Optional reusable image placement style.</param>
    /// <param name="linkUri">Optional absolute URI or catalog-base-relative URI for an image link annotation.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    /// <param name="alternativeText">Optional alternate text for meaningful generated images.</param>
    public PdfElementCompose Image(byte[] jpegBytes, double width, double height, PdfAlign? align, OfficeClipPath? clipPath, OfficeImageFit? fit, double? spacingBefore, double? spacingAfter, PdfImageStyle? style, string? linkUri, string? linkContents, string? alternativeText) { _doc.Image(jpegBytes, width, height, align, clipPath, fit, spacingBefore, spacingAfter, style, linkUri, linkContents, alternativeText); return this; }
}
