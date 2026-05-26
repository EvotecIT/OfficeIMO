using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Builder for nested elements used within item builders.</summary>
public class PdfElementCompose {
    private readonly PdfDoc _doc;
    internal PdfElementCompose(PdfDoc doc) { _doc = doc; }
    /// <summary>Adds invisible vertical space to the current flow.</summary>
    public PdfElementCompose Spacer(double height) { _doc.Spacer(height); return this; }
    /// <summary>Starts a new page from the current nested element flow.</summary>
    public PdfElementCompose PageBreak() { _doc.PageBreak(); return this; }
    /// <summary>Adds an H1 heading.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="style">Optional heading style.</param>
    /// <param name="linkUri">Optional absolute URI for a heading link annotation.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfElementCompose H1(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.H1(text, style: style, linkUri: linkUri, linkContents: linkContents); return this; }
    /// <summary>Adds an H1 heading with explicit alignment and color.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="align">Heading alignment.</param>
    /// <param name="color">Optional heading color.</param>
    /// <param name="linkUri">Optional absolute URI for a heading link annotation.</param>
    /// <param name="style">Optional heading style.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfElementCompose H1(string text, PdfAlign align, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null) { _doc.H1(text, align, color, linkUri, style, linkContents); return this; }
    /// <summary>Adds an H2 heading.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="style">Optional heading style.</param>
    /// <param name="linkUri">Optional absolute URI for a heading link annotation.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfElementCompose H2(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.H2(text, style: style, linkUri: linkUri, linkContents: linkContents); return this; }
    /// <summary>Adds an H2 heading with explicit alignment and color.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="align">Heading alignment.</param>
    /// <param name="color">Optional heading color.</param>
    /// <param name="linkUri">Optional absolute URI for a heading link annotation.</param>
    /// <param name="style">Optional heading style.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfElementCompose H2(string text, PdfAlign align, PdfColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null) { _doc.H2(text, align, color, linkUri, style, linkContents); return this; }
    /// <summary>Adds an H3 heading.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="style">Optional heading style.</param>
    /// <param name="linkUri">Optional absolute URI for a heading link annotation.</param>
    /// <param name="linkContents">Optional link annotation contents metadata.</param>
    public PdfElementCompose H3(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.H3(text, style: style, linkUri: linkUri, linkContents: linkContents); return this; }
    /// <summary>Adds an H3 heading with explicit alignment and color.</summary>
    /// <param name="text">Heading text.</param>
    /// <param name="align">Heading alignment.</param>
    /// <param name="color">Optional heading color.</param>
    /// <param name="linkUri">Optional absolute URI for a heading link annotation.</param>
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
    /// <summary>Adds a simple numbered list.</summary>
    public PdfElementCompose Numbered(System.Collections.Generic.IEnumerable<string> items, PdfAlign align = PdfAlign.Left, PdfColor? color = null, int startNumber = 1, PdfListStyle? style = null) { _doc.Numbered(items, align, color, startNumber, style); return this; }
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
    /// <param name="links">Per-cell absolute link URIs keyed by zero-based row and column.</param>
    /// <param name="align">Table alignment.</param>
    /// <param name="style">Optional table styling.</param>
    public PdfElementCompose TableWithLinks(System.Collections.Generic.IEnumerable<string[]> rows, System.Collections.Generic.Dictionary<(int Row, int Col), string> links, PdfAlign align = PdfAlign.Left, PdfTableStyle? style = null) { _doc.TableWithLinks(rows, links, align, style); return this; }
    /// <summary>Adds a named bookmark at the current nested element flow position.</summary>
    public PdfElementCompose Bookmark(string name) { _doc.Bookmark(name); return this; }
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
}
