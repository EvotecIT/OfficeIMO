#pragma warning disable RS0026 // The wrapper mirrors the PdfDocument operation overloads.
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Builds document flow at any supported nesting level. The same receiver is used for document
/// content, page content, rows, columns, containers, components, and deferred content.
/// </summary>
public sealed partial class PdfContentBuilder {
    private readonly PdfDocument _doc;
    internal PdfContentBuilder(PdfDocument doc) { _doc = doc; }
    /// <summary>Captures reusable footnote flow without inserting it into the body. Associate its handle with a styled text run.</summary>
    /// <param name="build">Static paragraphs, lists, panels, supported eager inline tables, images, rules, and spacers forming the footnote.</param>
    /// <param name="options">Placement and optional separator flow, snapshotted during this call.</param>
    /// <returns>An immutable handle belonging to this document.</returns>
    public PdfFootnoteDefinition DefineFootnote(Action<PdfContentBuilder> build, PdfFootnoteOptions? options = null) => _doc.DefineFootnote(build, options);
    internal PdfContentBuilder Heading(int level, string title, System.Action<PdfParagraph> build, OfficeAlignment align, global::OfficeIMO.OfficeColor? color, PdfHeadingStyle? style) {
        _doc.Heading(level, title, build, align, color, style);
        return this;
    }
    /// <summary>Returns this content receiver for readable grouping in fluent chains.</summary>
    public PdfContentBuilder Item() => this;
    /// <summary>Adds a logical group of content without introducing layout behavior.</summary>
    public PdfContentBuilder Item(System.Action<PdfContentBuilder> build) {
        Guard.NotNull(build, nameof(build));
        build(this);
        return this;
    }
    /// <summary>Adds a vertical content sequence using this same universal receiver.</summary>
    public PdfContentBuilder Column(System.Action<PdfContentBuilder> build) {
        Guard.NotNull(build, nameof(build));
        build(this);
        return this;
    }
    /// <summary>Adds a reusable typed component through the canonical flow engine.</summary>
    public PdfContentBuilder Component(IPdfComponent component, PdfFlowOptions? options = null, PdfLayoutPositionCapture? capture = null) { _doc.Component(component, options, capture); return this; }
    /// <summary>Adds a reusable page-aware component through the canonical deferred-flow engine.</summary>
    public PdfContentBuilder Component(IPdfContextComponent component, PdfFlowOptions? options = null, PdfLayoutPositionCapture? capture = null) { _doc.Component(component, options, capture); return this; }
    /// <summary>Adds a nested flow group with optional constraints and position capture.</summary>
    public PdfContentBuilder Flow(System.Action<PdfContentBuilder> build, PdfFlowOptions? options = null, PdfLayoutPositionCapture? capture = null) { _doc.Flow(build, options, capture); return this; }
    /// <summary>Adds replayable content materialized from the live page context during layout.</summary>
    public PdfContentBuilder Deferred(System.Func<PdfFlowContext, System.Action<PdfContentBuilder>> build, PdfFlowOptions? options = null, PdfLayoutPositionCapture? capture = null) { _doc.Deferred(build, options, capture); return this; }
    /// <summary>Adds a semantic section backed by an outline entry and named destination.</summary>
    public PdfContentBuilder Section(string title, System.Action<PdfContentBuilder> build, PdfSectionOptions? options = null) { _doc.Section(title, build, options); return this; }
    /// <summary>Groups content under an explicit tagged-PDF semantic role without changing its visual layout.</summary>
    public PdfContentBuilder Semantic(PdfSemanticRole role, System.Action<PdfContentBuilder> build, string? alternativeText = null) {
        Guard.NotNull(build, nameof(build));
        _doc.Semantic(role, build, alternativeText);
        return this;
    }
    /// <summary>Adds an automatically generated, internally linked table of contents.</summary>
    public PdfContentBuilder TableOfContents(PdfTableOfContentsOptions? options = null) { _doc.TableOfContents(options); return this; }
    /// <summary>Adds flow content controlled by a generated optional-content layer.</summary>
    public PdfContentBuilder Layer(string name, System.Action<PdfContentBuilder> build, PdfLayerOptions? options = null) { _doc.Layer(name, build, options); return this; }
    /// <summary>Flows common block content across equal-width columns in reading order.</summary>
    public PdfContentBuilder Columns(System.Action<PdfContentBuilder> build, PdfMultiColumnOptions? options = null) { _doc.Columns(build, options); return this; }
    /// <summary>Advances automatic multi-column flow to the next column.</summary>
    public PdfContentBuilder ColumnBreak() { _doc.AddColumnBreak(); return this; }
    /// <summary>Adds a row whose columns can use relative, fixed, automatic, or percentage sizing.</summary>
    public PdfContentBuilder Row(System.Action<PdfRowBuilder> build) { _doc.Row(build); return this; }
    /// <summary>Adds a deferred simple-text table whose rows are materialized in bounded batches.</summary>
    /// <remarks>Minimum first-page and final-page body row groups can enlarge a batch so its pagination constraints remain intact.</remarks>
    public PdfContentBuilder AddDeferredTable(System.Func<System.Collections.Generic.IEnumerable<string[]>> rowFactory, int batchSize = 256, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null) { _doc.AddDeferredTable(rowFactory, batchSize, align, style); return this; }
    /// <summary>Adds a deferred explicit-cell table whose rows are materialized in bounded batches.</summary>
    /// <remarks>Minimum first-page and final-page body row groups can enlarge a batch so its pagination constraints remain intact.</remarks>
    public PdfContentBuilder AddDeferredTable(System.Func<System.Collections.Generic.IEnumerable<PdfTableCell[]>> rowFactory, int batchSize = 256, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null) { _doc.AddDeferredTable(rowFactory, batchSize, align, style); return this; }
    /// <summary>Adds a two-column key/value table.</summary>
    public PdfContentBuilder AddKeyValueTable(System.Collections.Generic.IEnumerable<System.Collections.Generic.KeyValuePair<string, string?>> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null, bool includeHeader = false, string keyHeader = "Key", string valueHeader = "Value") { _doc.AddKeyValueTable(rows, align, style, includeHeader, keyHeader, valueHeader); return this; }
    /// <summary>Adds a two-column key/value table with rich row values.</summary>
    public PdfContentBuilder AddKeyValueTable(System.Collections.Generic.IEnumerable<PdfKeyValueRow> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null, bool includeHeader = false, string keyHeader = "Key", string valueHeader = "Value") { _doc.AddKeyValueTable(rows, align, style, includeHeader, keyHeader, valueHeader); return this; }
    /// <summary>Adds page-canvas drawing commands to the current flow.</summary>
    public PdfContentBuilder Canvas(System.Action<PdfPageCanvas> build) { _doc.Canvas(build); return this; }
    /// <summary>Starts a new page.</summary>
    public PdfContentBuilder PageBreak() { _doc.PageBreak(); return this; }
    /// <summary>Starts a new page, optionally preserving an empty current page for consecutive explicit breaks.</summary>
    public PdfContentBuilder PageBreak(bool preserveEmptyPage) { _doc.PageBreak(preserveEmptyPage); return this; }
    /// <summary>Adds invisible vertical space to the current flow.</summary>
    public PdfContentBuilder Spacer(global::OfficeIMO.OfficeLength height) { _doc.Spacer(height.ToPoints()); return this; }
    /// <summary>Adds a heading at an authored level from 1 through 9.</summary>
    /// <remarks>Levels 7 through 9 retain their outline depth and use custom tagged roles mapped to H6.</remarks>
    public PdfContentBuilder AddHeading(string text, int level = 1, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) {
        _doc.AddHeading(text, level, align, color, linkUri, style, linkContents, linkDestinationName);
        return this;
    }

    /// <summary>Adds a document title above the heading hierarchy. It uses the theme title size and stays out of the generated outline.</summary>
    public PdfContentBuilder Title(string text, OfficeAlignment align = OfficeAlignment.Left, OfficeColor? color = null, PdfHeadingStyle? style = null) { _doc.Title(text, align, color, style); return this; }
    /// <summary>Adds a subtitle paragraph in the theme accent colour.</summary>
    public PdfContentBuilder Subtitle(string text, OfficeAlignment align = OfficeAlignment.Left, OfficeColor? color = null) { _doc.Subtitle(text, align, color); return this; }
    /// <summary>Adds an H1 heading.</summary>
    public PdfContentBuilder H1(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null, string? linkDestinationName = null) { _doc.AddHeading(text, 1, style: style, linkUri: linkUri, linkContents: linkContents, linkDestinationName: linkDestinationName); return this; }
    /// <summary>Adds an H1 heading with explicit alignment and color.</summary>
    public PdfContentBuilder H1(string text, OfficeAlignment align, global::OfficeIMO.OfficeColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) { _doc.AddHeading(text, 1, align, color, linkUri, style, linkContents, linkDestinationName); return this; }
    /// <summary>Adds an H2 heading.</summary>
    public PdfContentBuilder H2(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null, string? linkDestinationName = null) { _doc.AddHeading(text, 2, style: style, linkUri: linkUri, linkContents: linkContents, linkDestinationName: linkDestinationName); return this; }
    /// <summary>Adds an H2 heading with explicit alignment and color.</summary>
    public PdfContentBuilder H2(string text, OfficeAlignment align, global::OfficeIMO.OfficeColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) { _doc.AddHeading(text, 2, align, color, linkUri, style, linkContents, linkDestinationName); return this; }
    /// <summary>Adds an H3 heading.</summary>
    public PdfContentBuilder H3(string text, PdfHeadingStyle? style = null, string? linkUri = null, string? linkContents = null, string? linkDestinationName = null) { _doc.AddHeading(text, 3, style: style, linkUri: linkUri, linkContents: linkContents, linkDestinationName: linkDestinationName); return this; }
    /// <summary>Adds an H3 heading with explicit alignment and color.</summary>
    public PdfContentBuilder H3(string text, OfficeAlignment align, global::OfficeIMO.OfficeColor? color = null, string? linkUri = null, PdfHeadingStyle? style = null, string? linkContents = null, string? linkDestinationName = null) { _doc.AddHeading(text, 3, align, color, linkUri, style, linkContents, linkDestinationName); return this; }
    /// <summary>Adds a paragraph with one text run. See <see cref="PdfDocument.AddParagraph(string, OfficeAlignment, OfficeColor?, PdfParagraphStyle?)"/>.</summary>
    public PdfContentBuilder AddParagraph(string text, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, PdfParagraphStyle? style = null) {
        Guard.NotNull(text, nameof(text));
        _doc.AddParagraph(text, align, color, style);
        return this;
    }
    /// <summary>Adds a paragraph built from styled text runs. See <see cref="PdfDocument.AddParagraph(System.Action{PdfParagraph}, OfficeAlignment, OfficeColor?, PdfParagraphStyle?)"/>.</summary>
    public PdfContentBuilder AddParagraph(System.Action<PdfParagraph> build, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, PdfParagraphStyle? style = null) { _doc.AddParagraph(build, align, color, style); return this; }
    /// <summary>Adds a bulleted list of plain items.</summary>
    public PdfContentBuilder AddList(System.Collections.Generic.IEnumerable<string> items, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, PdfListStyle? style = null) { _doc.AddList(items, align, color, style); return this; }
    /// <summary>Adds a bulleted list whose items can contain rich inline text runs.</summary>
    public PdfContentBuilder AddList(System.Collections.Generic.IEnumerable<PdfListItem> items, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, PdfListStyle? style = null) { _doc.AddList(items, align, color, style); return this; }
    /// <summary>Adds a numbered list of plain items.</summary>
    public PdfContentBuilder AddNumberedList(System.Collections.Generic.IEnumerable<string> items, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, int startNumber = 1, PdfListStyle? style = null) { _doc.AddNumberedList(items, align, color, startNumber, style); return this; }
    /// <summary>Adds a numbered list whose items can contain rich inline text runs.</summary>
    public PdfContentBuilder AddNumberedList(System.Collections.Generic.IEnumerable<PdfListItem> items, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null, int startNumber = 1, PdfListStyle? style = null) { _doc.AddNumberedList(items, align, color, startNumber, style); return this; }
    /// <summary>Adds a table of plain text rows with the document table style.</summary>
    public PdfContentBuilder AddTable(System.Collections.Generic.IEnumerable<string[]> rows) { _doc.AddTable(rows); return this; }
    /// <summary>Adds a table of explicit cells with the document table style.</summary>
    public PdfContentBuilder AddTable(System.Collections.Generic.IEnumerable<PdfTableCell[]> rows) { _doc.AddTable(rows); return this; }
    /// <summary>Adds a table of plain text rows.</summary>
    public PdfContentBuilder AddTable(System.Collections.Generic.IEnumerable<string[]> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null) { _doc.AddTable(rows, align, style); return this; }
    /// <summary>Adds a table of explicit cells, including spans, links, images, form fields and rich runs.</summary>
    public PdfContentBuilder AddTable(System.Collections.Generic.IEnumerable<PdfTableCell[]> rows, OfficeAlignment align = OfficeAlignment.Left, PdfTableStyle? style = null) { _doc.AddTable(rows, align, style); return this; }
    /// <summary>Adds a table from typed rows. See <see cref="PdfDocument.AddTable{T}(System.Collections.Generic.IEnumerable{T}, OfficeTableOptions?)"/>.</summary>
    public PdfContentBuilder AddTable<[System.Diagnostics.CodeAnalysis.DynamicallyAccessedMembers(System.Diagnostics.CodeAnalysis.DynamicallyAccessedMemberTypes.PublicProperties | System.Diagnostics.CodeAnalysis.DynamicallyAccessedMemberTypes.PublicFields)] T>(
        System.Collections.Generic.IEnumerable<T> rows, OfficeTableOptions? options = null) { _doc.AddTable(rows, options); return this; }
    /// <summary>Builds a visual and semantic envelope around nested flow content.</summary>
    /// <param name="build">Element layout, semantics, and content configuration.</param>
    public PdfContentBuilder Element(System.Action<PdfElementBuilder> build) {
        Guard.NotNull(build, nameof(build));
        var element = new PdfElementBuilder(_doc);
        build(element);
        element.Commit();
        return this;
    }
    /// <summary>Adds a horizontal rule.</summary>
    /// <param name="thickness">Line thickness (pt).</param>
    /// <param name="color">Optional color; inherited from the current default rule style when omitted.</param>
    /// <param name="spacingBefore">Top spacing (pt), inherited from the current default rule style when omitted.</param>
    /// <param name="spacingAfter">Bottom spacing (pt), inherited from the current default rule style when omitted.</param>
    /// <param name="style">Optional reusable rule style.</param>
    public PdfContentBuilder HR(global::OfficeIMO.OfficeLength? thickness = null, global::OfficeIMO.OfficeColor? color = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, PdfHorizontalRuleStyle? style = null) { _doc.HR((thickness?.ToPoints()), color, (spacingBefore?.ToPoints()), (spacingAfter?.ToPoints()), style); return this; }
    /// <summary>Adds a named bookmark at the current flow position.</summary>
    public PdfContentBuilder Bookmark(string name) { _doc.Bookmark(name); return this; }
    /// <summary>Adds a simple AcroForm text field at the current flow position.</summary>
    public PdfContentBuilder TextField(string name, OfficeLength? width = null, OfficeLength? height = null, string value = "", OfficeAlignment align = OfficeAlignment.Left, OfficeLength? fontSize = null, OfficeLength? spacingBefore = null, OfficeLength? spacingAfter = null, PdfFormFieldStyle? style = null) { _doc.TextField(name, width, height, value, align, fontSize, spacingBefore, spacingAfter, style); return this; }
    /// <summary>Adds a simple AcroForm check box at the current flow position.</summary>
    public PdfContentBuilder CheckBox(string name, bool isChecked = false, global::OfficeIMO.OfficeLength? size = null, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, string checkedValueName = "Yes", PdfFormFieldStyle? style = null) { _doc.CheckBox(name, isChecked, (size ?? global::OfficeIMO.OfficeLength.Points(14)).ToPoints(), align, (spacingBefore ?? global::OfficeIMO.OfficeLength.Points(0)).ToPoints(), (spacingAfter ?? global::OfficeIMO.OfficeLength.Points(6)).ToPoints(), checkedValueName, style); return this; }
    /// <summary>Adds a simple AcroForm choice field at the current flow position.</summary>
    public PdfContentBuilder ChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, global::OfficeIMO.OfficeLength? width = null, global::OfficeIMO.OfficeLength? height = null, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeLength? fontSize = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, bool isComboBox = true, PdfFormFieldStyle? style = null) { _doc.ChoiceField(name, options, value, (width ?? global::OfficeIMO.OfficeLength.Points(180)).ToPoints(), (height ?? global::OfficeIMO.OfficeLength.Points(22)).ToPoints(), align, (fontSize ?? global::OfficeIMO.OfficeLength.Points(10)).ToPoints(), (spacingBefore ?? global::OfficeIMO.OfficeLength.Points(0)).ToPoints(), (spacingAfter ?? global::OfficeIMO.OfficeLength.Points(6)).ToPoints(), isComboBox, style); return this; }
    /// <summary>Adds a simple AcroForm multi-select choice field at the current flow position.</summary>
    public PdfContentBuilder MultiSelectChoiceField(string name, System.Collections.Generic.IEnumerable<string> options, System.Collections.Generic.IEnumerable<string>? values = null, global::OfficeIMO.OfficeLength? width = null, global::OfficeIMO.OfficeLength? height = null, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeLength? fontSize = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, PdfFormFieldStyle? style = null) { _doc.MultiSelectChoiceField(name, options, values, (width ?? global::OfficeIMO.OfficeLength.Points(180)).ToPoints(), (height ?? global::OfficeIMO.OfficeLength.Points(72)).ToPoints(), align, (fontSize ?? global::OfficeIMO.OfficeLength.Points(10)).ToPoints(), (spacingBefore ?? global::OfficeIMO.OfficeLength.Points(0)).ToPoints(), (spacingAfter ?? global::OfficeIMO.OfficeLength.Points(6)).ToPoints(), style); return this; }
    /// <summary>Adds a simple AcroForm radio button group at the current flow position.</summary>
    public PdfContentBuilder RadioButtonGroup(string name, System.Collections.Generic.IEnumerable<string> options, string? value = null, global::OfficeIMO.OfficeLength? size = null, global::OfficeIMO.OfficeLength? gap = null, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, PdfFormFieldStyle? style = null) { _doc.RadioButtonGroup(name, options, value, (size ?? global::OfficeIMO.OfficeLength.Points(14)).ToPoints(), (gap ?? global::OfficeIMO.OfficeLength.Points(6)).ToPoints(), align, (spacingBefore ?? global::OfficeIMO.OfficeLength.Points(0)).ToPoints(), (spacingAfter ?? global::OfficeIMO.OfficeLength.Points(6)).ToPoints(), style); return this; }
    /// <summary>Adds a PDF text annotation at the current flow position.</summary>
    public PdfContentBuilder TextAnnotation(string contents, global::OfficeIMO.OfficeLength? width = null, global::OfficeIMO.OfficeLength? height = null, OfficeAlignment? align = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, PdfTextAnnotationIcon icon = PdfTextAnnotationIcon.Comment, global::OfficeIMO.OfficeColor? color = null, bool open = false) { _doc.TextAnnotation(contents, (width ?? global::OfficeIMO.OfficeLength.Points(18D)).ToPoints(), (height ?? global::OfficeIMO.OfficeLength.Points(18D)).ToPoints(), align, (spacingBefore?.ToPoints()), (spacingAfter?.ToPoints()), icon, color, open); return this; }
    /// <summary>Adds a PDF free-text annotation at the current flow position.</summary>
    public PdfContentBuilder FreeTextAnnotation(string contents, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, OfficeAlignment? align = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, global::OfficeIMO.OfficeLength? fontSize = null, global::OfficeIMO.OfficeColor? textColor = null, global::OfficeIMO.OfficeColor? borderColor = null, global::OfficeIMO.OfficeLength? borderWidth = null, global::OfficeIMO.OfficeColor? fillColor = null, OfficeAlignment textAlign = OfficeAlignment.Left, global::OfficeIMO.OfficeLength? padding = null, OfficeLength? lineHeight = null) { _doc.FreeTextAnnotation(contents, width.ToPoints(), height.ToPoints(), align, (spacingBefore?.ToPoints()), (spacingAfter?.ToPoints()), (fontSize ?? global::OfficeIMO.OfficeLength.Points(10D)).ToPoints(), textColor, borderColor, (borderWidth ?? global::OfficeIMO.OfficeLength.Points(1D)).ToPoints(), fillColor, textAlign, (padding ?? global::OfficeIMO.OfficeLength.Points(3D)).ToPoints(), lineHeight?.ToPoints()); return this; }
    /// <summary>Adds a PDF highlight annotation rectangle at the current flow position.</summary>
    public PdfContentBuilder HighlightAnnotation(string contents, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, OfficeAlignment? align = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, global::OfficeIMO.OfficeColor? color = null) { _doc.HighlightAnnotation(contents, width.ToPoints(), height.ToPoints(), align, (spacingBefore?.ToPoints()), (spacingAfter?.ToPoints()), color); return this; }
    /// <summary>Adds a shared OfficeIMO.Drawing shape.</summary>
    public PdfContentBuilder Shape(OfficeShape shape, OfficeAlignment? align = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Shape(shape, align, (spacingBefore?.ToPoints()), (spacingAfter?.ToPoints()), style, linkUri, linkContents); return this; }
    /// <summary>Adds a shared OfficeIMO.Drawing scene.</summary>
    public PdfContentBuilder Drawing(OfficeDrawing drawing, OfficeAlignment? align = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Drawing(drawing, align, (spacingBefore?.ToPoints()), (spacingAfter?.ToPoints()), style, linkUri, linkContents); return this; }
    /// <summary>Adds a line vector shape.</summary>
    public PdfContentBuilder Line(global::OfficeIMO.OfficeLength x1, global::OfficeIMO.OfficeLength y1, global::OfficeIMO.OfficeLength x2, global::OfficeIMO.OfficeLength y2, global::OfficeIMO.OfficeColor? strokeColor = null, global::OfficeIMO.OfficeLength? strokeWidth = null, OfficeAlignment? align = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Line(x1.ToPoints(), y1.ToPoints(), x2.ToPoints(), y2.ToPoints(), strokeColor, (strokeWidth ?? global::OfficeIMO.OfficeLength.Points(1)).ToPoints(), align, (spacingBefore?.ToPoints()), (spacingAfter?.ToPoints()), strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a rectangle vector shape.</summary>
    public PdfContentBuilder Rectangle(global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, global::OfficeIMO.OfficeColor? strokeColor = null, global::OfficeIMO.OfficeLength? strokeWidth = null, global::OfficeIMO.OfficeColor? fillColor = null, OfficeAlignment? align = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Rectangle(width.ToPoints(), height.ToPoints(), strokeColor, (strokeWidth ?? global::OfficeIMO.OfficeLength.Points(1)).ToPoints(), fillColor, align, (spacingBefore?.ToPoints()), (spacingAfter?.ToPoints()), strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a rounded rectangle vector shape.</summary>
    public PdfContentBuilder RoundedRectangle(global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, global::OfficeIMO.OfficeLength cornerRadius, global::OfficeIMO.OfficeColor? strokeColor = null, global::OfficeIMO.OfficeLength? strokeWidth = null, global::OfficeIMO.OfficeColor? fillColor = null, OfficeAlignment? align = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.RoundedRectangle(width.ToPoints(), height.ToPoints(), cornerRadius.ToPoints(), strokeColor, (strokeWidth ?? global::OfficeIMO.OfficeLength.Points(1)).ToPoints(), fillColor, align, (spacingBefore?.ToPoints()), (spacingAfter?.ToPoints()), strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds an ellipse vector shape.</summary>
    public PdfContentBuilder Ellipse(global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, global::OfficeIMO.OfficeColor? strokeColor = null, global::OfficeIMO.OfficeLength? strokeWidth = null, global::OfficeIMO.OfficeColor? fillColor = null, OfficeAlignment? align = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Ellipse(width.ToPoints(), height.ToPoints(), strokeColor, (strokeWidth ?? global::OfficeIMO.OfficeLength.Points(1)).ToPoints(), fillColor, align, (spacingBefore?.ToPoints()), (spacingAfter?.ToPoints()), strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a polygon vector shape.</summary>
    public PdfContentBuilder Polygon(System.Collections.Generic.IEnumerable<OfficePoint> points, global::OfficeIMO.OfficeColor? strokeColor = null, global::OfficeIMO.OfficeLength? strokeWidth = null, global::OfficeIMO.OfficeColor? fillColor = null, OfficeAlignment? align = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Polygon(points, strokeColor, (strokeWidth ?? global::OfficeIMO.OfficeLength.Points(1)).ToPoints(), fillColor, align, (spacingBefore?.ToPoints()), (spacingAfter?.ToPoints()), strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a freeform path vector shape.</summary>
    public PdfContentBuilder Path(System.Collections.Generic.IEnumerable<OfficePathCommand> commands, global::OfficeIMO.OfficeColor? strokeColor = null, global::OfficeIMO.OfficeLength? strokeWidth = null, global::OfficeIMO.OfficeColor? fillColor = null, OfficeAlignment? align = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) { _doc.Path(commands, strokeColor, (strokeWidth ?? global::OfficeIMO.OfficeLength.Points(1)).ToPoints(), fillColor, align, (spacingBefore?.ToPoints()), (spacingAfter?.ToPoints()), strokeDashStyle, strokeLineCap, strokeLineJoin, style, linkUri, linkContents); return this; }
    /// <summary>Adds a paragraph inside a styled panel.</summary>
    /// <param name="build">Paragraph content builder.</param>
    /// <param name="style">Panel style (padding, background, border, etc.).</param>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="color">Optional text colour.</param>
    public PdfContentBuilder AddPanelParagraph(System.Action<PdfParagraph> build, PdfPanelStyle? style = null, OfficeAlignment align = OfficeAlignment.Left, global::OfficeIMO.OfficeColor? color = null) { _doc.PanelParagraph(build, style, align, color); return this; }
    /// <summary>Adds a styled element while preserving its paragraphs, tables, links, and semantic structure.</summary>
    /// <param name="build">Panel content builder.</param>
    /// <param name="style">Panel style (padding, background, border, etc.).</param>
    public PdfContentBuilder Panel(System.Action<PdfContentBuilder> build, PdfPanelStyle? style = null) { _doc.Panel(build, style); return this; }
    /// <summary>Adds a meaningful image from supported image bytes with alternative text.</summary>
    public PdfContentBuilder AddImage(byte[] imageBytes, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? altText) { _doc.AddImage(imageBytes, width, height, altText); return this; }

    /// <summary>Adds a raster image supported by OfficeIMO.Drawing. See <see cref="PdfDocument.AddImage(byte[], OfficeLength, OfficeLength, OfficeAlignment?, OfficeClipPath?, OfficeImageFit?, OfficeLength?, OfficeLength?, PdfImageStyle?, string?, string?, string?)"/>.</summary>
    public PdfContentBuilder AddImage(byte[] imageBytes, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, OfficeAlignment? align = null, OfficeClipPath? clipPath = null, OfficeImageFit? fit = null, global::OfficeIMO.OfficeLength? spacingBefore = null, global::OfficeIMO.OfficeLength? spacingAfter = null, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null, string? altText = null) {
        _doc.AddImage(imageBytes, width, height, align, clipPath, fit, spacingBefore, spacingAfter, style, linkUri, linkContents, altText);
        return this;
    }
}
