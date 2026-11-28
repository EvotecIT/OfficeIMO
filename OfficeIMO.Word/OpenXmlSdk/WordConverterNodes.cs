using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Word;

/// <summary>
/// The elements and parts of a document as <see cref="OfficeXmlNode"/> handles, for the converter packages
/// (<c>OfficeIMO.Word.Pdf</c>, <c>.Html</c>, <c>.Rtf</c>, <c>.Markdown</c>). Converters read documents through these and
/// <see cref="OpenXml.WordXml"/> only, so they name no Open XML SDK type.
/// </summary>
/// <remarks>
/// While the model keeps SDK elements, every handle here wraps the SDK element, which opens the SDK package as the
/// model's own members do; handles from different accessors then compare equal when they are the same element. This
/// file is the converters' only path to the SDK: it shrinks as the model moves to OfficeIMO's own tree.
/// </remarks>
internal static class WordConverterNodes {
    // ---- Document parts -------------------------------------------------------------------------------------------

    /// <summary>The <c>w:document</c> element.</summary>
    internal static OfficeXmlNode DocumentNode(this WordDocument document) => OfficeXmlNode.FromSdk(document._document);

    /// <summary>The <c>w:body</c> element, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode BodyNode(this WordDocument document) => OfficeXmlNode.FromSdk(document._document?.Body);

    private static MainDocumentPart? Main(WordDocument? document) => document?._wordprocessingDocument?.MainDocumentPart;

    /// <summary>The <c>w:settings</c> root, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode SettingsNode(this WordDocument? document) =>
        OfficeXmlNode.FromSdk(Main(document)?.DocumentSettingsPart?.Settings);

    /// <summary>
    /// The <c>w:styles</c> root, or <see cref="OfficeXmlNode.None"/>. Reading it loads the whole styles part; style
    /// lookups go through <see cref="WordStyleCatalog"/>, which streams it, and <see cref="HasStylesPart"/> tells
    /// whether there is one.
    /// </summary>
    internal static OfficeXmlNode StylesNode(this WordDocument? document) =>
        OfficeXmlNode.FromSdk(Main(document)?.StyleDefinitionsPart?.Styles);

    /// <summary>Whether the document has a styles part, without loading it.</summary>
    internal static bool HasStylesPart(this WordDocument? document) => Main(document)?.StyleDefinitionsPart != null;

    /// <summary>The <c>w:numbering</c> root, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode NumberingNode(this WordDocument? document) =>
        OfficeXmlNode.FromSdk(Main(document)?.NumberingDefinitionsPart?.Numbering);

    /// <summary>The <c>w:footnotes</c> root, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode FootnotesNode(this WordDocument? document) =>
        OfficeXmlNode.FromSdk(Main(document)?.FootnotesPart?.Footnotes);

    /// <summary>The <c>w:endnotes</c> root, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode EndnotesNode(this WordDocument? document) =>
        OfficeXmlNode.FromSdk(Main(document)?.EndnotesPart?.Endnotes);

    /// <summary>The <c>w:comments</c> root, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode CommentsNode(this WordDocument? document) =>
        OfficeXmlNode.FromSdk(Main(document)?.WordprocessingCommentsPart?.Comments);

    /// <summary>The <c>w:fonts</c> root of the font table, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode FontTableNode(this WordDocument? document) =>
        OfficeXmlNode.FromSdk(Main(document)?.FontTablePart?.Fonts);

    /// <summary>The <c>a:theme</c> root, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode ThemeNode(this WordDocument? document) =>
        OfficeXmlNode.FromSdk(Main(document)?.ThemePart?.Theme);

    // ---- Parts and relationships ----------------------------------------------------------------------------------

    /// <summary>The root element of the part that holds <paramref name="element"/>.</summary>
    internal static OfficeXmlNode PartRoot(this OfficeXmlNode element) =>
        OfficeXmlNode.FromSdk((PartOf(element))?.RootElement);

    /// <summary>The root element of the part <paramref name="relationshipId"/> names from the part that holds <paramref name="element"/>.</summary>
    internal static OfficeXmlNode RelatedPartRoot(this OfficeXmlNode element, string? relationshipId) =>
        OfficeXmlNode.FromSdk(RelatedPart(element, relationshipId)?.RootElement);

    /// <summary>The root element of the part <paramref name="relationshipId"/> names from the main document part.</summary>
    internal static OfficeXmlNode RelatedPartRoot(this WordDocument document, string? relationshipId) =>
        OfficeXmlNode.FromSdk(TryGetPart(Main(document), relationshipId)?.RootElement);

    /// <summary>
    /// Whether <paramref name="relationshipId"/> names an image part from the part that holds <paramref name="element"/>;
    /// <paramref name="bytes"/> is its content, or null when it holds more than <paramref name="maxBytes"/>.
    /// </summary>
    internal static bool TryGetRelatedImagePartBytes(this OfficeXmlNode element, string? relationshipId, long maxBytes, out byte[]? bytes) {
        bytes = null;
        if (RelatedPart(element, relationshipId) is not ImagePart part) return false;
        bytes = ReadPartBytes(part, maxBytes);
        return true;
    }

    private static byte[]? ReadPartBytes(OpenXmlPart? part, long maxBytes) {
        if (part == null || maxBytes < 0L) return null;
        using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
        if (stream.CanSeek && stream.Length > maxBytes) return null;
        using var copy = new MemoryStream();
        byte[] buffer = new byte[81920];
        long total = 0L;
        int read;
        while ((read = stream.Read(buffer, 0, buffer.Length)) > 0) {
            total += read;
            if (total > maxBytes) return null;
            copy.Write(buffer, 0, read);
        }
        return copy.ToArray();
    }

    private static OpenXmlPart? PartOf(OfficeXmlNode element) =>
        element.SdkElement is { } sdk ? OfficeSdkXml.PartOf(sdk) as OpenXmlPart : null;

    private static OpenXmlPart? RelatedPart(OfficeXmlNode element, string? relationshipId) =>
        TryGetPart(PartOf(element), relationshipId);

    private static OpenXmlPart? TryGetPart(OpenXmlPart? part, string? relationshipId) {
        if (part == null || string.IsNullOrEmpty(relationshipId)) return null;
        return part.TryGetPartById(relationshipId!, out OpenXmlPart? related) ? related : null;
    }

    // ---- Model objects --------------------------------------------------------------------------------------------

    /// <summary>The paragraph's <c>w:p</c>.</summary>
    internal static OfficeXmlNode Node(this WordParagraph paragraph) => OfficeXmlNode.FromSdk(paragraph._paragraph);

    /// <summary>The paragraph's <c>w:pPr</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode ParagraphPropertiesNode(this WordParagraph paragraph) =>
        OfficeXmlNode.FromSdk(paragraph._paragraphProperties);

    /// <summary>The run this paragraph object stands for, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode RunNode(this WordParagraph paragraph) => OfficeXmlNode.FromSdk(paragraph._run);

    /// <summary>The run properties this paragraph object reads, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode RunPropertiesNode(this WordParagraph paragraph) => OfficeXmlNode.FromSdk(paragraph._runProperties);

    /// <summary>The hyperlink this paragraph object stands in, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode HyperlinkNode(this WordParagraph paragraph) => OfficeXmlNode.FromSdk(paragraph._hyperlink);

    /// <summary>The table's <c>w:tbl</c>.</summary>
    internal static OfficeXmlNode Node(this WordTable table) => OfficeXmlNode.FromSdk(table._table);

    /// <summary>The table's <c>w:tblPr</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode PropertiesNode(this WordTable table) => OfficeXmlNode.FromSdk(table._tableProperties);

    /// <summary>The row's <c>w:tr</c>.</summary>
    internal static OfficeXmlNode Node(this WordTableRow row) => OfficeXmlNode.FromSdk(row._tableRow);

    /// <summary>The cell's <c>w:tc</c>.</summary>
    internal static OfficeXmlNode Node(this WordTableCell cell) => OfficeXmlNode.FromSdk(cell._tableCell);

    /// <summary>The cell's <c>w:tcPr</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode PropertiesNode(this WordTableCell cell) => OfficeXmlNode.FromSdk(cell._tableCellProperties);

    /// <summary>The section's <c>w:sectPr</c>.</summary>
    internal static OfficeXmlNode Node(this WordSection section) => OfficeXmlNode.FromSdk(section._sectionProperties);

    /// <summary>The <c>w:hdr</c> or <c>w:ftr</c> root of the header or footer, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode Node(this WordHeaderFooter headerFooter) =>
        OfficeXmlNode.FromSdk((object?)headerFooter._header ?? headerFooter._footer);

    /// <summary>The image's <c>w:drawing</c>, or <see cref="OfficeXmlNode.None"/> for a VML image.</summary>
    internal static OfficeXmlNode DrawingNode(this WordImage image) => OfficeXmlNode.FromSdk(image._Image);

    /// <summary>The image's VML <c>v:shape</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode VmlShapeNode(this WordImage image) => OfficeXmlNode.FromSdk(image._vmlShape);

    /// <summary>The element that places the image in the story: its VML shape or its drawing.</summary>
    internal static OfficeXmlNode OccurrenceNode(this WordImage image) => OfficeXmlNode.FromSdk(image.OccurrenceElement);

    /// <summary>The chart's <c>w:drawing</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode DrawingNode(this WordChart chart) => OfficeXmlNode.FromSdk(chart.Drawing);

    /// <summary>The <c>c:chartSpace</c> root of the chart part, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode ChartSpaceNode(this WordChart chart) => OfficeXmlNode.FromSdk(chart.ChartPart?.ChartSpace);

    /// <summary>The SmartArt's <c>w:drawing</c>.</summary>
    internal static OfficeXmlNode DrawingNode(this WordSmartArt smartArt) => OfficeXmlNode.FromSdk(smartArt._drawing);

    /// <summary>The text box's <c>w:drawing</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode DrawingNode(this WordTextBox textBox) => OfficeXmlNode.FromSdk(textBox.Drawing);

    /// <summary>The text box's <c>wp:anchor</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode AnchorNode(this WordTextBox textBox) => OfficeXmlNode.FromSdk(textBox.Anchor);

    /// <summary>The text box's <c>w:txbxContent</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode ContentNode(this WordTextBox textBox) => OfficeXmlNode.FromSdk(textBox.Content);

    /// <summary>The text box shape's <c>wps:spPr</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode ShapePropertiesNode(this WordTextBox textBox) => OfficeXmlNode.FromSdk(textBox.DrawingShapeProperties);

    /// <summary>The text box shape's <c>wps:bodyPr</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode TextBodyPropertiesNode(this WordTextBox textBox) => OfficeXmlNode.FromSdk(textBox.DrawingTextBodyProperties);

    /// <summary>The shape's <c>w:drawing</c>, or <see cref="OfficeXmlNode.None"/> for a VML shape.</summary>
    internal static OfficeXmlNode DrawingNode(this WordShape shape) => OfficeXmlNode.FromSdk(shape._drawing);

    /// <summary>The shape's <c>wps:wsp</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode WpsShapeNode(this WordShape shape) => OfficeXmlNode.FromSdk(shape._wpsShape);

    /// <summary>The run that holds the shape, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode RunNode(this WordShape shape) => OfficeXmlNode.FromSdk(shape.Run);

    /// <summary>The content control's <c>w:sdt</c> block.</summary>
    internal static OfficeXmlNode SdtBlockNode(this WordStructuredDocumentTag tag) => OfficeXmlNode.FromSdk(tag.SdtBlock);

    /// <summary>The cover page's <c>w:sdt</c> block.</summary>
    internal static OfficeXmlNode SdtBlockNode(this WordCoverPage coverPage) => OfficeXmlNode.FromSdk(coverPage.SdtBlock);

    /// <summary>The table of contents' <c>w:sdt</c> block.</summary>
    internal static OfficeXmlNode SdtBlockNode(this WordTableOfContent tableOfContent) => OfficeXmlNode.FromSdk(tableOfContent.SdtBlock);

    /// <summary>The shape's VML <c>v:line</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode VmlLineNode(this WordShape shape) => OfficeXmlNode.FromSdk(shape._line);

    /// <summary>The shape's VML <c>v:polyline</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode VmlPolygonNode(this WordShape shape) => OfficeXmlNode.FromSdk(shape._polygon);

    /// <summary>The shape's VML <c>v:oval</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode VmlOvalNode(this WordShape shape) => OfficeXmlNode.FromSdk(shape._ellipse);

    /// <summary>The shape's VML <c>v:roundrect</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode VmlRoundRectangleNode(this WordShape shape) => OfficeXmlNode.FromSdk(shape._roundRectangle);

    /// <summary>The URI of the chart's part, as text, or null.</summary>
    internal static string? ChartPartUri(this WordChart chart) => chart.ChartPart?.Uri?.ToString();

    /// <summary>The inline content control (<c>w:sdt</c>) this paragraph object stands for, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode SdtRunNode(this WordParagraph paragraph) => OfficeXmlNode.FromSdk(paragraph._stdRun);

    /// <summary>The run properties the hyperlink's run carries, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode RunPropertiesNode(this WordHyperLink hyperlink) => OfficeXmlNode.FromSdk(hyperlink._runProperties);

    /// <summary>The style's <c>w:style</c> element, parsed when first asked for.</summary>
    internal static OfficeXmlNode ElementNode(this WordStyleDefinitionsReader.StyleDefinition style) => OfficeXmlNode.FromSdk(style.Element);

    /// <summary>The <c>w:docDefaults</c> of the style catalog, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode DocumentDefaultsNode(this WordStyleCatalog catalog) => OfficeXmlNode.FromSdk(catalog.DocumentDefaults);

    /// <summary>The <c>w:docDefaults</c> of the document's styles, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode DocumentDefaultsNode(this WordDocument document) =>
        OfficeXmlNode.FromSdk(WordStyleCatalog.For(document).DocumentDefaults);

    /// <summary>The <c>w:tabs</c> of the paragraph's numbering level, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode LevelTabStopsNode(this WordDocumentTraversal.ListInfo info) => OfficeXmlNode.FromSdk(info.LevelTabStops);

    /// <summary>The element the field was read from.</summary>
    internal static OfficeXmlNode SourceNode(this WordFieldInventory.FieldInfoCandidate field) => OfficeXmlNode.FromSdk(field.SourceElement);

    // ---- Run content ----------------------------------------------------------------------------------------------

    /// <summary>The run's direct content with the active markup-compatibility branch selected.</summary>
    internal static IEnumerable<OfficeXmlNode> EffectiveRunContentNodes(this WordParagraph paragraph) {
        foreach (OpenXmlElement child in paragraph.EnumerateEffectiveRunContent()) yield return OfficeXmlNode.FromSdk(child);
    }

    /// <summary>The source children a run view projects, or null when it shows the whole source run.</summary>
    internal static IReadOnlyList<OfficeXmlNode>? VisibleRunSourceChildNodes(this WordParagraph paragraph) =>
        paragraph._visibleRunSourceChildren is { } children ? ToNodes(children) : null;

    /// <summary>
    /// The run content a run view shows: its visible source children when it was narrowed to them, otherwise every
    /// child of its run.
    /// </summary>
    internal static IEnumerable<OfficeXmlNode> ProjectedRunChildNodes(this WordParagraph run) {
        IEnumerable<OpenXmlElement> children = run._visibleRunSourceChildren ?? run._run!.ChildElements;
        foreach (OpenXmlElement child in children) yield return OfficeXmlNode.FromSdk(child);
    }

    /// <summary>The branch of an <c>mc:AlternateContent</c> Word selects, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode SelectAlternateContentBranch(OfficeXmlNode alternateContent) =>
        alternateContent.SdkElement is AlternateContent alternate
            ? OfficeXmlNode.FromSdk(WordAlternateContentResolver.SelectBranch(alternate))
            : default;

    /// <summary>The element's visible text, as <see cref="WordParagraph.ReadVisibleText(OpenXmlElement)"/>.</summary>
    internal static string VisibleText(this OfficeXmlNode element) => WordParagraph.ReadVisibleText((OpenXmlElement)element.SdkElement!);

    /// <summary>The element's visible text with non-text breaks replaced, as <c>WordParagraph.ReadVisibleText</c>.</summary>
    internal static string VisibleText(this OfficeXmlNode element, char nonTextBreakReplacement) =>
        WordParagraph.ReadVisibleText((OpenXmlElement)element.SdkElement!, nonTextBreakReplacement);

    /// <summary>The text of the Office Math below <paramref name="element"/>.</summary>
    internal static string MathText(OfficeXmlNode element) => WordMath.GetText((OpenXmlElement)element.SdkElement!);

    // ---- Model views over nodes -----------------------------------------------------------------------------------

    /// <summary>
    /// Whether the model has an object for <paramref name="element"/>. Walks that match elements by name also reach
    /// elements the package holds untyped, inside unknown or ignorable markup (a <c>w:p</c> in a vendor wrapper, for
    /// example); the model can't wrap those (<see cref="ParagraphView"/> and the other views need a typed element), and
    /// its own typed walks never return them. Elements of OfficeIMO's own tree always qualify.
    /// </summary>
    internal static bool IsModelElement(this OfficeXmlNode element) =>
        element.Exists && element.SdkElement is not OpenXmlUnknownElement;

    /// <summary>
    /// The <c>w:</c> descendants named <paramref name="localName"/> the model can wrap, in document order: the elements
    /// the model's own typed walks return.
    /// </summary>
    internal static IEnumerable<OfficeXmlNode> ModelDescendants(this OfficeXmlNode element, string localName) =>
        OpenXml.WordXml.DescendantElements(element, OpenXml.WordXml.W, localName).Where(IsModelElement);

    /// <summary>A paragraph object over a <c>w:p</c>, standing for its bookmark when <paramref name="bookmarkStart"/> is given.</summary>
    internal static WordParagraph ParagraphView(this WordDocument document, OfficeXmlNode paragraph, OfficeXmlNode bookmarkStart = default) {
        var view = new WordParagraph(document, (W.Paragraph)paragraph.SdkElement!);
        if (bookmarkStart.SdkElement is W.BookmarkStart start) view._bookmarkStart = start;
        return view;
    }

    /// <summary>A paragraph object over a new, detached, empty <c>w:p</c>.</summary>
    internal static WordParagraph CreateDetachedParagraphView(this WordDocument document) => new(document, new W.Paragraph());

    /// <summary>
    /// A run object over <paramref name="run"/> in <paramref name="paragraph"/>, optionally inside a hyperlink and
    /// showing <paramref name="visibleRun"/> projected from <paramref name="visibleRunSourceChildren"/>.
    /// </summary>
    internal static WordParagraph RunView(this WordDocument document, OfficeXmlNode paragraph, OfficeXmlNode run,
        OfficeXmlNode hyperlink = default, OfficeXmlNode visibleRun = default,
        IReadOnlyList<OfficeXmlNode>? visibleRunSourceChildren = null) =>
        new(document, (W.Paragraph)paragraph.SdkElement!, (W.Run)run.SdkElement!) {
            _hyperlink = (W.Hyperlink?)hyperlink.SdkElement,
            _visibleRun = (W.Run?)visibleRun.SdkElement,
            _visibleRunSourceChildren = visibleRunSourceChildren == null ? null : ToElements(visibleRunSourceChildren)
        };

    /// <summary>A paragraph object over an inline content control in <paramref name="paragraph"/>.</summary>
    internal static WordParagraph SdtRunView(this WordDocument document, OfficeXmlNode paragraph, OfficeXmlNode sdtRun) =>
        new(document, (W.Paragraph)paragraph.SdkElement!, (W.SdtRun)sdtRun.SdkElement!);

    /// <summary>
    /// A paragraph object for the paragraph mark: a detached run that carries a copy of the mark's run properties,
    /// so the mark's font, size and emphasis resolve as an ordinary run's.
    /// </summary>
    internal static WordParagraph ParagraphMarkRunView(this WordParagraph paragraph) {
        var properties = new W.RunProperties();
        W.ParagraphMarkRunProperties? mark = paragraph._paragraph.ParagraphProperties?.ParagraphMarkRunProperties;
        if (mark != null) {
            foreach (OpenXmlElement property in mark.ChildElements) properties.Append(property.CloneNode(true));
        }
        return new WordParagraph(paragraph._document, paragraph._paragraph, new W.Run(properties));
    }

    /// <summary>
    /// A detached paragraph with one run of <paramref name="text"/> in style <paramref name="styleId"/>, its built-in style
    /// reference localized as the document stores it, for measuring the formatting a style gives.
    /// </summary>
    internal static WordParagraph CreateDetachedStyledParagraph(this WordDocument document, string styleId, string text) {
        var paragraph = new W.Paragraph(
            new W.ParagraphProperties(new W.ParagraphStyleId { Val = styleId }),
            new W.Run(new W.Text(text)));
        document.LocalizeBuiltInStyleReferences(paragraph);
        return new WordParagraph(document, paragraph, paragraph.GetFirstChild<W.Run>()!);
    }

    /// <summary>The canonical paragraph views of a <c>w:p</c>, bookmarks included.</summary>
    internal static List<WordParagraph> ConvertParagraphNodeToWordParagraphs(this WordDocument document, OfficeXmlNode paragraph) =>
        WordSection.ConvertParagraphToWordParagraphs(document, (W.Paragraph)paragraph.SdkElement!);

    /// <summary>A table object for a <c>w:tbl</c>.</summary>
    internal static WordTable CreateTable(WordDocument document, OfficeXmlNode table) => new(document, (W.Table)table.SdkElement!);

    /// <summary>A read-only table view over a <c>w:tbl</c> that builds no child objects up front.</summary>
    internal static WordTable ReadOnlyTableView(this WordDocument document, OfficeXmlNode table) =>
        new(document, (W.Table)table.SdkElement!, initializeChildren: false, readOnly: true);

    /// <summary>A shape object for a drawing in a run of a paragraph.</summary>
    internal static WordShape CreateShape(WordDocument document, OfficeXmlNode paragraph, OfficeXmlNode run, OfficeXmlNode drawing) =>
        new(document, (W.Paragraph)paragraph.SdkElement!, (W.Run)run.SdkElement!, drawing.SdkElement as W.Drawing);

    /// <summary>
    /// A shape for one <c>wps:wsp</c> of a DrawingML group, read in the run and drawing that hold the group, as
    /// <see cref="WordShape"/>'s own constructor builds it.
    /// </summary>
    internal static WordShape CreateGroupChildShape(WordParagraph run, OfficeXmlNode drawing, OfficeXmlNode wpsShape) =>
        new(run._document, run._paragraph!, run._run!, (W.Drawing?)drawing.SdkElement) {
            _wpsShape = (Wps.WordprocessingShape?)wpsShape.SdkElement
        };

    /// <summary>
    /// A tab stop of <paramref name="paragraph"/> over a detached copy of a <c>w:tab</c>. An alignment or leader token
    /// the schema doesn't define is left out of the copy, so it reads as missing.
    /// </summary>
    internal static WordTabStop DetachedTabStop(WordParagraph paragraph, OfficeXmlNode tabStop) {
        var copy = (W.TabStop)((W.TabStop)tabStop.SdkElement!).CloneNode(true);
        if (copy.Val is { HasValue: false }) copy.Val = null;
        if (copy.Leader is { HasValue: false }) copy.Leader = null;
        return new WordTabStop(paragraph, copy);
    }

    // ---- Detached elements ----------------------------------------------------------------------------------------

    /// <summary>A new, detached, empty <c>w:p</c>.</summary>
    internal static OfficeXmlNode CreateDetachedParagraph() => OfficeXmlNode.FromSdk(new W.Paragraph());

    /// <summary>A new, detached, empty <c>w:r</c>.</summary>
    internal static OfficeXmlNode CreateDetachedRun() => OfficeXmlNode.FromSdk(new W.Run());

    /// <summary>A new, detached <c>w:t</c> holding <paramref name="text"/> with spaces preserved.</summary>
    internal static OfficeXmlNode CreateDetachedPreservedText(string text) =>
        OfficeXmlNode.FromSdk(new W.Text(text) { Space = SpaceProcessingModeValues.Preserve });

    /// <summary>A detached copy of the element with its attributes and no children, as the SDK's <c>CloneNode(false)</c>.</summary>
    internal static OfficeXmlNode CloneShallow(this OfficeXmlNode element) {
        if (element.SdkElement is OpenXmlElement sdk) return OfficeXmlNode.FromSdk(sdk.CloneNode(false));
        OfficeXmlNode copy = element.Clone();
        for (OfficeXmlNode child = copy.FirstChild; !child.IsNone; child = copy.FirstChild) child.Remove();
        return copy;
    }

    /// <summary>Lets detached paragraph copies resolve the list information of the paragraphs they were copied from.</summary>
    internal static IDisposable UseListParagraphAliases(WordDocument document, IReadOnlyDictionary<OfficeXmlNode, OfficeXmlNode> copiesToSources) {
        var aliases = new Dictionary<W.Paragraph, W.Paragraph>(copiesToSources.Count);
        foreach (KeyValuePair<OfficeXmlNode, OfficeXmlNode> alias in copiesToSources) {
            // Paragraphs the model can't wrap have no list information to resolve.
            if (alias.Key.SdkElement is W.Paragraph copy && alias.Value.SdkElement is W.Paragraph source) aliases[copy] = source;
        }
        return WordDocumentTraversal.UseListParagraphAliases(document, aliases);
    }

    /// <summary>
    /// Lets the detached copy <paramref name="root"/> resolve its relationships against the part that holds
    /// <paramref name="source"/>.
    /// </summary>
    internal static IDisposable UseDetachedOwner(OfficeXmlNode root, WordDocument document, OfficeXmlNode source) =>
        WordPartOwnership.UseDetachedOwner((OpenXmlElement)root.SdkElement!,
            WordPartOwnership.Resolve(document, (OpenXmlElement)source.SdkElement!));

    // ---- Complex-field visibility ---------------------------------------------------------------------------------

    /// <summary>The complex-field visibility at the start of <paramref name="paragraph"/>.</summary>
    internal static WordComplexFieldRunVisibility FieldVisibilityForParagraph(OfficeXmlNode paragraph) =>
        WordComplexFieldRunVisibility.ForParagraph((W.Paragraph)paragraph.SdkElement!);

    /// <summary>Whether no enclosing field hides <paramref name="element"/>.</summary>
    internal static bool IsFieldVisibleBefore(OfficeXmlNode element) =>
        WordComplexFieldRunVisibility.IsVisibleBefore((OpenXmlElement)element.SdkElement!);

    /// <summary>The <c>w:fldChar</c> that began the innermost open field, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode BeginMarkerNode(this WordComplexFieldRunVisibility visibility) =>
        OfficeXmlNode.FromSdk(visibility.CurrentState?.BeginMarker);

    /// <summary>
    /// The visible form of <paramref name="run"/>: the run itself, a projection holding only its visible children, or
    /// <see cref="OfficeXmlNode.None"/> when a field hides it.
    /// </summary>
    internal static OfficeXmlNode GetVisibleRunNode(this WordComplexFieldRunVisibility visibility, OfficeXmlNode run,
        out IReadOnlyList<OfficeXmlNode>? visibleSourceChildren) {
        W.Run? visible = visibility.GetVisibleRun((W.Run)run.SdkElement!, out IReadOnlyList<OpenXmlElement>? children);
        visibleSourceChildren = children == null ? null : ToNodes(children);
        return OfficeXmlNode.FromSdk(visible);
    }

    /// <summary>Advances the field state past the <c>w:r</c> <paramref name="run"/>, observing its field characters.</summary>
    internal static void ObserveRun(this WordComplexFieldRunVisibility visibility, OfficeXmlNode run) =>
        visibility.GetVisibleRun((W.Run)run.SdkElement!);

    /// <summary>Advances the field state over the runs below <paramref name="container"/>.</summary>
    internal static void ObserveDescendantRuns(this WordComplexFieldRunVisibility visibility, OfficeXmlNode container) =>
        visibility.ObserveDescendantRuns((OpenXmlElement)container.SdkElement!);

    /// <summary>Carries the source paragraph's field state onto the two fragments split at <paramref name="boundary"/>.</summary>
    internal static void PreserveFragmentContext(OfficeXmlNode source, OfficeXmlNode before, OfficeXmlNode after, OfficeXmlNode boundary) =>
        WordComplexFieldRunVisibility.PreserveFragmentContext((W.Paragraph)source.SdkElement!, (W.Paragraph)before.SdkElement!,
            (W.Paragraph)after.SdkElement!, (OpenXmlElement)boundary.SdkElement!);

    // ---- Equations and content controls ---------------------------------------------------------------------------

    /// <summary>The visible equations of <paramref name="paragraph"/>, in child order.</summary>
    internal static IReadOnlyList<WordEquationOccurrence> EquationOccurrences(this WordDocument document, OfficeXmlNode paragraph) =>
        WordEquation.GetOccurrences(document, (W.Paragraph)paragraph.SdkElement!);

    /// <summary>Whether the element backs the equation, as <see cref="WordEquation.IsBackingElement"/>.</summary>
    internal static bool IsBackingNode(this WordEquation equation, OfficeXmlNode element) =>
        element.SdkElement is OpenXmlElement sdk && equation.IsBackingElement(sdk);

    /// <summary>The visible text and equation segments of <paramref name="container"/>.</summary>
    internal static IReadOnlyList<WordEquationContentSegment> VisibleEquationContentSegments(OfficeXmlNode container,
        IReadOnlyList<WordEquationOccurrence> occurrences, Func<OfficeXmlNode, bool>? includeElement = null) =>
        WordEquation.GetVisibleContentSegments((OpenXmlElement)container.SdkElement!, occurrences,
            includeElement == null ? null : element => includeElement(OfficeXmlNode.FromSdk(element)));

    /// <summary>The run artifact (break, note reference, ...) the equation segment stands for, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode ArtifactNode(this WordEquationContentSegment segment) => OfficeXmlNode.FromSdk(segment.ArtifactElement);

    /// <summary>The run object the equation segment's text comes from, or <paramref name="paragraph"/>.</summary>
    internal static WordParagraph CreateSourceParagraph(this WordEquationContentSegment segment, WordParagraph paragraph) =>
        segment.CreateSourceParagraph(paragraph._document, paragraph._paragraph, paragraph);

    /// <summary>The run object the segment's text comes from in <paramref name="paragraph"/>, or <paramref name="fallback"/>.</summary>
    internal static WordParagraph CreateSourceParagraph(this WordEquationContentSegment segment, WordDocument document,
        OfficeXmlNode paragraph, WordParagraph fallback) =>
        segment.CreateSourceParagraph(document, (W.Paragraph)paragraph.SdkElement!, fallback);

    /// <summary>The item texts of a repeating-section content control.</summary>
    internal static IReadOnlyList<string> ReadRepeatingSectionItemTexts(OfficeXmlNode repeatingSection) =>
        WordRepeatingSection.ReadItemTexts((W.SdtRun)repeatingSection.SdkElement!);

    /// <summary>The elements a block content control shows, as <see cref="WordDocumentTraversal.GetBlockContentElements"/> projects them.</summary>
    internal static IReadOnlyList<WordElement> GetBlockContentElements(WordDocument document, OfficeXmlNode block,
        Func<OfficeXmlNode, bool>? includeParagraph = null) =>
        WordDocumentTraversal.GetBlockContentElements(document, (W.SdtBlock?)block.SdkElement,
            includeParagraph == null ? null : paragraph => includeParagraph(OfficeXmlNode.FromSdk(paragraph)));

    // ---- Notes ----------------------------------------------------------------------------------------------------

    /// <summary>The ordinary footnote with this identifier, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode FindFootnoteNode(this WordDocument? document, long? referenceId) =>
        OfficeXmlNode.FromSdk(WordNoteResolver.FindFootnote(Main(document), referenceId));

    /// <summary>The ordinary endnote with this identifier, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode FindEndnoteNode(this WordDocument? document, long? referenceId) =>
        OfficeXmlNode.FromSdk(WordNoteResolver.FindEndnote(Main(document), referenceId));

    /// <summary>Whether a <c>w:footnote</c> or <c>w:endnote</c> is an ordinary note rather than a separator story.</summary>
    internal static bool IsOrdinaryNote(OfficeXmlNode note) =>
        WordNoteResolver.IsOrdinary((W.FootnoteEndnoteType)note.SdkElement!);

    /// <summary>
    /// The detached default separator (or continuation separator) story Word uses when the part registers none: one
    /// single-spaced paragraph holding the mark. The package is not changed.
    /// </summary>
    internal static OfficeXmlNode CreateImplicitNoteSeparatorStory(bool endnote, bool continuation) {
        W.FootnoteEndnoteType story = endnote ? new W.Endnote() : new W.Footnote();
        story.Type = continuation ? W.FootnoteEndnoteValues.ContinuationSeparator : W.FootnoteEndnoteValues.Separator;
        story.Id = continuation ? 0 : -1;
        OpenXmlElement mark = continuation ? new W.ContinuationSeparatorMark() : new W.SeparatorMark();
        story.Append(new W.Paragraph(
            new W.ParagraphProperties(new W.SpacingBetweenLines { Line = "240", LineRule = W.LineSpacingRuleValues.Auto }),
            new W.Run(mark)));
        return OfficeXmlNode.FromSdk(story);
    }

    // ---- Formatting -----------------------------------------------------------------------------------------------

    /// <summary>A <c>w:color</c> as an RGB hex value, theme references first; null when the element is missing.</summary>
    internal static string? ResolveColorHex(OfficeXmlNode color, WordDocument? document) =>
        WordRunFormatting.ResolveColorHex(color.SdkElement as W.Color, document);

    /// <summary>A <c>w:shd</c> fill resolved against the document theme, as RGB hex, or null.</summary>
    internal static string? ResolveShadingFillColorHex(OfficeXmlNode shading, WordDocument document) =>
        WordRunFormatting.ResolveShadingFillColorHex(shading.SdkElement as W.Shading, document);

    /// <summary>A border's color resolved against the document theme, as RGB hex, or null.</summary>
    internal static string? ResolveBorderColorHex(OfficeXmlNode border, WordDocument document) =>
        WordRunFormatting.ResolveBorderColorHex(border.SdkElement as W.BorderType, document);

    /// <summary>The shading of the nearest table cell holding the paragraph, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode ContainingTableCellShadingNode(this WordParagraph paragraph) =>
        OfficeXmlNode.FromSdk(WordRunFormatting.ResolveContainingTableCellShading(paragraph));

    /// <summary>The cell's resolved <c>w:shd</c>, as <see cref="WordRunFormatting.ResolveTableCellShading(WordTableCell)"/>.</summary>
    internal static OfficeXmlNode ResolveTableCellShading(WordTableCell cell) =>
        OfficeXmlNode.FromSdk(WordRunFormatting.ResolveTableCellShading(cell));

    /// <summary>The table's resolved <c>w:shd</c>, as <see cref="WordRunFormatting.ResolveTableShading(WordTable)"/>.</summary>
    internal static OfficeXmlNode ResolveTableShading(WordTable table) =>
        OfficeXmlNode.FromSdk(WordRunFormatting.ResolveTableShading(table));

    /// <summary>The highlight token of run properties, as <see cref="WordRunFormatting.ReadRunHighlight"/>.</summary>
    internal static string? ReadRunHighlight(OfficeXmlNode properties) =>
        EnumToken(WordRunFormatting.ReadRunHighlight(properties.SdkElement as OpenXmlCompositeElement));

    /// <summary>The run's highlight token from its properties, styles, the table style and the document defaults.</summary>
    internal static string? ResolveRunHighlightToken(WordParagraph paragraph, string? tableStyleHighlight = null) =>
        EnumToken(WordRunFormatting.ResolveRunHighlight(paragraph,
            tableStyleHighlight == null ? null : new W.HighlightColorValues(tableStyleHighlight)));

    /// <summary>The RGB hex value of a highlight token; none and unknown tokens have no paint.</summary>
    internal static string? HighlightColorHexOfToken(string? highlight) =>
        WordRunFormatting.GetHighlightColorHex(highlight == null ? null : new W.HighlightColorValues(highlight));

    /// <summary>A <c>w:tblHeader</c> flag as the table row reads it: a missing value is on, explicit off forms are off.</summary>
    internal static bool ReadTableHeaderOnOff(OfficeXmlNode tableHeader) =>
        WordTableRow.ReadOnOff((W.TableHeader)tableHeader.SdkElement!);

    /// <summary>The resolved type token of a preferred table width, as <see cref="WordTableWidthNative.ResolveType"/>.</summary>
    internal static string? ResolveTableWidthType(OfficeXmlNode width) =>
        EnumToken(WordTableWidthNative.ResolveType(width.SdkElement as W.TableWidthType));

    /// <summary>The resolved type token of a cell width, as <see cref="WordTableWidthNative.ResolveCellType"/>.</summary>
    internal static string? ResolveTableCellWidthType(OfficeXmlNode width) =>
        EnumToken(WordTableWidthNative.ResolveCellType(width.SdkElement as W.TableCellWidth));

    /// <summary>The table alignment token of a <c>w:jc</c> in table properties, as <see cref="WordTableAlignmentResolver.Read"/>.</summary>
    internal static string? ReadTableAlignment(OfficeXmlNode justification, bool rightToLeft) =>
        EnumToken(WordTableAlignmentResolver.Read(justification.SdkElement as W.TableJustification, rightToLeft));

    /// <summary>Whether table properties lay the table out right to left, as <see cref="WordTableAlignmentResolver.IsRightToLeft"/>.</summary>
    internal static bool IsTableRightToLeft(OfficeXmlNode tableProperties) =>
        WordTableAlignmentResolver.IsRightToLeft(tableProperties.SdkElement as W.TableProperties);

    // ---- Drawings and packages ------------------------------------------------------------------------------------

    /// <summary>Reads a <c>w:drawing</c>'s placement and size, as <see cref="WordDrawingLayoutReader.TryRead"/>.</summary>
    internal static bool TryReadDrawingLayout(OfficeXmlNode drawing, out WordDrawingLayoutInfo layout) =>
        WordDrawingLayoutReader.TryRead(drawing.SdkElement as W.Drawing, out layout);

    /// <summary>The VML shape's style declarations, as <see cref="ImageShapeStyleHelper.GetStyle"/>.</summary>
    internal static Dictionary<string, string> VmlShapeStyle(OfficeXmlNode shape) => ImageShapeStyleHelper.GetStyle((V.Shape)shape.SdkElement!);

    /// <summary>
    /// The <c>a:theme</c> root of the main document part of the package that holds <paramref name="element"/>, or
    /// <see cref="OfficeXmlNode.None"/> for a detached element or a package without a theme.
    /// </summary>
    internal static OfficeXmlNode PackageThemeNode(this OfficeXmlNode element) =>
        PartOf(element)?.OpenXmlPackage is WordprocessingDocument document
            ? OfficeXmlNode.FromSdk(document.MainDocumentPart?.ThemePart?.Theme)
            : default;

    /// <summary>The <c>w:hdr</c> roots of the header parts, then the <c>w:ftr</c> roots of the footer parts.</summary>
    internal static IReadOnlyList<OfficeXmlNode> HeaderFooterRootNodes(this WordDocument document) {
        var roots = new List<OfficeXmlNode>();
        MainDocumentPart? main = Main(document);
        if (main == null) return roots;
        foreach (HeaderPart part in main.HeaderParts) {
            if (part.Header is { } header) roots.Add(OfficeXmlNode.FromSdk(header));
        }
        foreach (FooterPart part in main.FooterParts) {
            if (part.Footer is { } footer) roots.Add(OfficeXmlNode.FromSdk(footer));
        }
        return roots;
    }

    /// <summary>
    /// The target of an external hyperlink relationship <paramref name="relationshipId"/> from the part that holds the
    /// paragraph, or null when it is missing or internal.
    /// </summary>
    internal static Uri? ExternalHyperlinkTarget(this WordParagraph paragraph, string? relationshipId) {
        HyperlinkRelationship? relationship = WordHyperLink.FindHyperlinkRelationship(
            WordHyperLink.ResolveRelationshipPart(paragraph._document, paragraph._paragraph), relationshipId);
        return relationship is { IsExternal: true } ? relationship.Uri : null;
    }

    /// <summary>Writes a copy of the document's package to <paramref name="destination"/>, unchanged and without saving.</summary>
    internal static void ClonePackage(this WordDocument document, Stream destination) {
        using (document._wordprocessingDocument.Clone(destination, true)) {
        }
    }

    /// <summary>Whether a package read failed on a part relationship layout the package reader rejects.</summary>
    internal static bool IsPackageReadException(Exception exception) => exception is OpenXmlPackageException;

    // ---- Helpers --------------------------------------------------------------------------------------------------

    private static string? EnumToken<T>(T? value) where T : struct, IEnumValue => value.HasValue ? value.Value.Value : null;

    private static List<OfficeXmlNode> ToNodes(IReadOnlyList<OpenXmlElement> elements) {
        var nodes = new List<OfficeXmlNode>(elements.Count);
        foreach (OpenXmlElement element in elements) nodes.Add(OfficeXmlNode.FromSdk(element));
        return nodes;
    }

    private static List<OpenXmlElement> ToElements(IReadOnlyList<OfficeXmlNode> nodes) {
        var elements = new List<OpenXmlElement>(nodes.Count);
        foreach (OfficeXmlNode node in nodes) elements.Add((OpenXmlElement)node.SdkElement!);
        return elements;
    }

    // ---- Word.Html (read) ---

    /// <summary>
    /// A style and its ancestors of the <c>w:type</c> token <paramref name="type"/> (<c>paragraph</c>, <c>character</c>,
    /// <c>table</c>, <c>numbering</c>), as <see cref="WordStyleCatalog.EnumerateChain"/> reads them.
    /// </summary>
    internal static IEnumerable<OfficeXmlNode> StyleChainNodes(this WordStyleCatalog catalog, string? styleId, string type) {
        foreach (W.Style style in catalog.EnumerateChain(styleId, StyleType(type))) yield return OfficeXmlNode.FromSdk(style);
    }

    private static W.StyleValues StyleType(string type) => type switch {
        "paragraph" => W.StyleValues.Paragraph,
        "character" => W.StyleValues.Character,
        "table" => W.StyleValues.Table,
        "numbering" => W.StyleValues.Numbering,
        _ => new W.StyleValues(type)
    };

    /// <summary>The first <c>w:style</c> with this identifier, as <see cref="WordStyleCatalog.FindById"/>.</summary>
    internal static OfficeXmlNode FindStyleNodeById(this WordStyleCatalog catalog, string? styleId) =>
        OfficeXmlNode.FromSdk(catalog.FindById(styleId));

    /// <summary>Whether the table's look turns first-row formatting on, as <see cref="WordTable.HasFirstRowFormatting"/>.</summary>
    internal static bool HasFirstRowFormatting(OfficeXmlNode table) => WordTable.HasFirstRowFormatting((W.Table)table.SdkElement!);

    /// <summary>The check box's <c>w:sdt</c>.</summary>
    internal static OfficeXmlNode SdtRunNode(this WordCheckBox checkBox) => OfficeXmlNode.FromSdk(checkBox._sdtRun);

    /// <summary>The drop-down list's <c>w:sdt</c>.</summary>
    internal static OfficeXmlNode SdtRunNode(this WordDropDownList dropDownList) => OfficeXmlNode.FromSdk(dropDownList._sdtRun);

    /// <summary>The combo box's <c>w:sdt</c>.</summary>
    internal static OfficeXmlNode SdtRunNode(this WordComboBox comboBox) => OfficeXmlNode.FromSdk(comboBox._sdtRun);

    /// <summary>The date picker's <c>w:sdt</c>.</summary>
    internal static OfficeXmlNode SdtRunNode(this WordDatePicker datePicker) => OfficeXmlNode.FromSdk(datePicker._sdtRun);

    /// <summary>The content control's inline or block <c>w:sdt</c>, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode SdtElementNode(this WordStructuredDocumentTag tag) => OfficeXmlNode.FromSdk(tag.SdtElement);

    /// <summary>The cell's resolved border on <paramref name="side"/>, as <see cref="WordRunFormatting.ResolveTableCellBorder(WordTableCell, WordTableBorderSide)"/>.</summary>
    internal static OfficeXmlNode ResolveTableCellBorderNode(WordTableCell cell, WordTableBorderSide side) =>
        OfficeXmlNode.FromSdk(WordRunFormatting.ResolveTableCellBorder(cell, side));

    /// <summary>The cell's resolved diagonal border, as <see cref="WordRunFormatting.ResolveTableCellDiagonalBorder"/>.</summary>
    internal static OfficeXmlNode ResolveTableCellDiagonalBorderNode(WordTableCell cell, bool descending) =>
        OfficeXmlNode.FromSdk(WordRunFormatting.ResolveTableCellDiagonalBorder(cell, descending));

    /// <summary>The footnote a <c>w:footnoteReference</c> in <paramref name="run"/>'s paragraph names, read from a detached run holding a copy of the reference.</summary>
    internal static WordFootNote FootNoteForReference(this WordParagraph run, OfficeXmlNode footnoteReference) =>
        new(run._document, run._paragraph, new W.Run((W.FootnoteReference)((W.FootnoteReference)footnoteReference.SdkElement!).CloneNode(true)));

    /// <summary>The endnote a <c>w:endnoteReference</c> in <paramref name="run"/>'s paragraph names, read from a detached run holding a copy of the reference.</summary>
    internal static WordEndNote EndNoteForReference(this WordParagraph run, OfficeXmlNode endnoteReference) =>
        new(run._document, run._paragraph, new W.Run((W.EndnoteReference)((W.EndnoteReference)endnoteReference.SdkElement!).CloneNode(true)));

    /// <summary>
    /// The part roots the HTML export inspects, in the order it reads them: custom properties (when asked), the field
    /// roots (body, headers, footers, notes, comments), chart parts of the stories and the charts they hold, comments,
    /// styles, numbering, theme, font table, settings and web settings. A styles part the streaming reader can count
    /// without loading has no root and carries its element count instead.
    /// </summary>
    internal static IEnumerable<(OfficeXmlNode Root, string PartUri, long StreamedElements)> HtmlExportRoots(
        this WordDocument document, bool includeCustomProperties) {
        MainDocumentPart? mainPart = Main(document);
        if (mainPart == null) yield break;
        CustomFilePropertiesPart? customPropertiesPart = includeCustomProperties ? document._wordprocessingDocument.CustomFilePropertiesPart : null;
        if (customPropertiesPart?.Properties is OpenXmlElement customProperties) {
            yield return (OfficeXmlNode.FromSdk(customProperties), customPropertiesPart.Uri.ToString(), 0);
        }

        foreach (WordFieldInventory.FieldRoot root in WordFieldInventory.EnumerateFieldRoots(mainPart)) {
            yield return (OfficeXmlNode.FromSdk(root.Root), root.PartUri, 0);
        }
        var chartUris = new HashSet<Uri>();
        IEnumerable<OpenXmlPart> stories = new OpenXmlPart[] { mainPart }
            .Concat(mainPart.HeaderParts).Concat(mainPart.FooterParts);
        if (mainPart.FootnotesPart != null) stories = stories.Append(mainPart.FootnotesPart);
        if (mainPart.EndnotesPart != null) stories = stories.Append(mainPart.EndnotesPart);
        foreach (var story in stories) {
            foreach (var chart in story.Parts.Select(relationship => relationship.OpenXmlPart)
                .Where(part => part.ContentType.IndexOf("chart", StringComparison.OrdinalIgnoreCase) >= 0)) {
                if (chartUris.Add(chart.Uri) && chart.RootElement is OpenXmlElement chartRoot)
                    yield return (OfficeXmlNode.FromSdk(chartRoot), chart.Uri.ToString(), 0);
                foreach (var presentation in chart.Parts.Select(relationship => relationship.OpenXmlPart)
                    .Where(part => part.ContentType.IndexOf("chart", StringComparison.OrdinalIgnoreCase) >= 0)) {
                    if (chartUris.Add(presentation.Uri) && presentation.RootElement is OpenXmlElement presentationRoot)
                        yield return (OfficeXmlNode.FromSdk(presentationRoot), presentation.Uri.ToString(), 0);
                }
            }
        }
        if (mainPart.WordprocessingCommentsPart?.Comments is OpenXmlElement comments) {
            yield return (OfficeXmlNode.FromSdk(comments), mainPart.WordprocessingCommentsPart.Uri.ToString(), 0);
        }
        foreach (StylesPart? stylesPart in new StylesPart?[] { mainPart.StyleDefinitionsPart, mainPart.StylesWithEffectsPart }) {
            if (stylesPart == null) continue;
            if (WordStyleDefinitionsReader.TryCountElements(stylesPart) is long count) {
                yield return (default, stylesPart.Uri.ToString(), count);
            } else if (stylesPart.Styles is OpenXmlElement styles) {
                yield return (OfficeXmlNode.FromSdk(styles), stylesPart.Uri.ToString(), 0);
            }
        }
        if (mainPart.NumberingDefinitionsPart?.Numbering is OpenXmlElement numbering) {
            yield return (OfficeXmlNode.FromSdk(numbering), mainPart.NumberingDefinitionsPart.Uri.ToString(), 0);
        }
        if (mainPart.ThemePart?.Theme is OpenXmlElement theme) {
            yield return (OfficeXmlNode.FromSdk(theme), mainPart.ThemePart.Uri.ToString(), 0);
        }
        if (mainPart.FontTablePart?.Fonts is OpenXmlElement fonts) {
            yield return (OfficeXmlNode.FromSdk(fonts), mainPart.FontTablePart.Uri.ToString(), 0);
        }
        if (mainPart.DocumentSettingsPart?.Settings is OpenXmlElement settings) {
            yield return (OfficeXmlNode.FromSdk(settings), mainPart.DocumentSettingsPart.Uri.ToString(), 0);
        }
        if (mainPart.WebSettingsPart?.WebSettings is OpenXmlElement webSettings) {
            yield return (OfficeXmlNode.FromSdk(webSettings), mainPart.WebSettingsPart.Uri.ToString(), 0);
        }
    }

    /// <summary>The text of an element the SDK reads as a leaf text element (<c>w:t</c>, <c>w:instrText</c>, <c>m:t</c>, ...), or null for any other element.</summary>
    internal static string? LeafText(this OfficeXmlNode element) => element.SdkElement is OpenXmlLeafTextElement text ? text.Text : null;

    /// <summary>Whether the package holds a digital signature origin part.</summary>
    internal static bool HasDigitalSignatureOriginPart(this WordDocument document) =>
        document._wordprocessingDocument.DigitalSignatureOriginPart != null;

    /// <summary>Puts <paramref name="body"/> in place of the document's <c>w:body</c>.</summary>
    internal static void SetBodyNode(this WordDocument document, OfficeXmlNode body) => document._document.Body = (W.Body)body.SdkElement!;

    /// <summary>
    /// Makes <paramref name="replacement"/> the root of the part <paramref name="currentRoot"/> (a <c>w:hdr</c>,
    /// <c>w:ftr</c>, <c>w:footnotes</c>, <c>w:endnotes</c> or <c>w:comments</c>) belongs to. False when it belongs to none.
    /// </summary>
    internal static bool ReplacePartRootNode(OfficeXmlNode currentRoot, OfficeXmlNode replacement) {
        switch (currentRoot.SdkElement) {
            case W.Header { HeaderPart: { } part }: part.Header = (W.Header)replacement.SdkElement!; return true;
            case W.Footer { FooterPart: { } part }: part.Footer = (W.Footer)replacement.SdkElement!; return true;
            case W.Footnotes { FootnotesPart: { } part }: part.Footnotes = (W.Footnotes)replacement.SdkElement!; return true;
            case W.Endnotes { EndnotesPart: { } part }: part.Endnotes = (W.Endnotes)replacement.SdkElement!; return true;
            case W.Comments { WordprocessingCommentsPart: { } part }: part.Comments = (W.Comments)replacement.SdkElement!; return true;
            default: return false;
        }
    }

    /// <summary>Makes the header or footer object read <paramref name="root"/> (a <c>w:hdr</c> or <c>w:ftr</c>).</summary>
    internal static void SetRootNode(this WordHeaderFooter headerFooter, OfficeXmlNode root) {
        if (root.SdkElement is W.Header header) headerFooter._header = header;
        else headerFooter._footer = (W.Footer)root.SdkElement!;
    }

    /// <summary>
    /// Opens the Word package in <paramref name="package"/> read-only and validates it with the validator's default
    /// file format version.
    /// </summary>
    internal static List<OfficeOpenXmlValidationError> ValidatePackage(Stream package) {
        using WordprocessingDocument document = WordprocessingDocument.Open(package, false);
        var errors = new List<OfficeOpenXmlValidationError>();
        foreach (DocumentFormat.OpenXml.Validation.ValidationErrorInfo error in new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(document)) {
            errors.Add(error.ToOfficeValidationError());
        }
        return errors;
    }

    /// <summary>
    /// The identifier the main document part gives the header part (<paramref name="header"/>) or footer part of the
    /// story, or null when the story has no such part or the main part doesn't reference it.
    /// </summary>
    internal static string? MainPartRelationshipId(this WordHeaderFooter story, bool header) {
        MainDocumentPart? main = Main(story.Document);
        OpenXmlPart? part = header ? story._header?.HeaderPart : story._footer?.FooterPart;
        if (main == null || part == null || !main.Parts.Any(relationship => ReferenceEquals(relationship.OpenXmlPart, part))) return null;
        return main.GetIdOfPart(part);
    }

    /// <summary>The run properties the run's formatting reads, its content control's included, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode ScopedRunPropertiesNode(this WordParagraph run) => OfficeXmlNode.FromSdk(run.ScopedRunProperties);

    /// <summary>A hyperlink object over the <c>w:hyperlink</c> <paramref name="hyperlink"/> in <paramref name="paragraph"/>.</summary>
    internal static WordHyperLink HyperlinkView(this WordDocument document, OfficeXmlNode paragraph, OfficeXmlNode hyperlink) =>
        new(document, (W.Paragraph)paragraph.SdkElement!, (W.Hyperlink)hyperlink.SdkElement!);

    /// <summary>Whether the <c>w:drawing</c> holds a chart, as <see cref="WordParagraph.DrawingOwnsChart"/>.</summary>
    internal static bool DrawingOwnsChart(OfficeXmlNode drawing) => WordParagraph.DrawingOwnsChart((W.Drawing)drawing.SdkElement!);

    /// <summary>A chart object over the chart <c>w:drawing</c> <paramref name="drawing"/> in <paramref name="run"/>.</summary>
    internal static WordChart ChartView(WordParagraph run, OfficeXmlNode drawing) =>
        new(run._document, run, (W.Drawing)drawing.SdkElement!);

    /// <summary>An image object over the picture <c>w:drawing</c> <paramref name="drawing"/>.</summary>
    internal static WordImage ImageView(this WordDocument document, OfficeXmlNode drawing) =>
        new(document, (W.Drawing)drawing.SdkElement!);

    /// <summary>An image object over the VML <c>v:shape</c> <paramref name="shape"/> in <paramref name="run"/>.</summary>
    internal static WordImage VmlImageView(WordParagraph run, OfficeXmlNode shape) =>
        new(run._document, run._paragraph, run._run!, (V.Shape)shape.SdkElement!);

    /// <summary>Whether the paragraph child carries visible equation content, as <see cref="WordEquation.IsVisibleEquationContentContainer"/>.</summary>
    internal static bool IsVisibleEquationContentContainer(OfficeXmlNode element) =>
        WordEquation.IsVisibleEquationContentContainer((OpenXmlElement)element.SdkElement!);

    /// <summary>The direct child of <paramref name="paragraph"/> that holds <paramref name="element"/>, as <see cref="WordEquation.GetDirectParagraphChild"/>.</summary>
    internal static OfficeXmlNode GetDirectParagraphChild(OfficeXmlNode paragraph, OfficeXmlNode element) =>
        OfficeXmlNode.FromSdk(WordEquation.GetDirectParagraphChild((W.Paragraph)paragraph.SdkElement!, element.SdkElement as OpenXmlElement));

    /// <summary>The hyperlink's <c>w:hyperlink</c>.</summary>
    internal static OfficeXmlNode Node(this WordHyperLink hyperlink) => OfficeXmlNode.FromSdk(hyperlink._hyperlink);

    // ---- Word.Html (build) ---

    /// <summary>Attaches an in-memory annotation object to the element; it isn't written to the document.</summary>
    internal static void AddAnnotation(this OfficeXmlNode element, object annotation) =>
        ((OpenXmlElement)element.SdkElement!).AddAnnotation(annotation);

    /// <summary>The element's first annotation of type <typeparamref name="T"/>, or null.</summary>
    internal static T? Annotation<T>(this OfficeXmlNode element) where T : class =>
        (element.SdkElement as OpenXmlElement)?.Annotation<T>();

    /// <summary>Removes the element's annotations of type <typeparamref name="T"/>.</summary>
    internal static void RemoveAnnotations<T>(this OfficeXmlNode element) where T : class =>
        (element.SdkElement as OpenXmlElement)?.RemoveAnnotations<T>();

    /// <summary>The Word elements for generated block elements (paragraphs, tables, content controls), in order.</summary>
    internal static IReadOnlyList<WordElement> WrapBlockElements(WordDocument document, IEnumerable<OfficeXmlNode> blocks,
        CancellationToken cancellationToken = default) =>
        WordSection.WrapBlockElements(document, blocks.Select(block => (OpenXmlElement)block.SdkElement!), cancellationToken);

    /// <summary>The last body block and last paragraph of the section, when the body layout matches it cheaply.</summary>
    internal static bool TryGetLastBlockNodes(this WordSection section, out OfficeXmlNode lastBlock, out OfficeXmlNode lastParagraph) {
        bool found = section.TryGetLastBlocks(out OpenXmlElement? block, out W.Paragraph? paragraph);
        lastBlock = OfficeXmlNode.FromSdk(block);
        lastParagraph = OfficeXmlNode.FromSdk(paragraph);
        return found;
    }

    /// <summary>
    /// The body paragraphs after <paramref name="lastParagraph"/> and tables after <paramref name="lastBlock"/> in the
    /// section, as <see cref="WordSection.TryGetBlocksAfter"/> collects them.
    /// </summary>
    internal static bool TryGetBlockNodesAfter(this WordSection section, OfficeXmlNode lastBlock, OfficeXmlNode lastParagraph,
        List<OfficeXmlNode> paragraphs, List<OfficeXmlNode> tables) {
        var sdkParagraphs = new List<W.Paragraph>();
        var sdkTables = new List<W.Table>();
        bool found = section.TryGetBlocksAfter((OpenXmlElement?)lastBlock.SdkElement, (W.Paragraph?)lastParagraph.SdkElement,
            sdkParagraphs, sdkTables);
        paragraphs.Clear();
        tables.Clear();
        foreach (W.Paragraph paragraph in sdkParagraphs) paragraphs.Add(OfficeXmlNode.FromSdk(paragraph));
        foreach (W.Table table in sdkTables) tables.Add(OfficeXmlNode.FromSdk(table));
        return found;
    }


    /// <summary>The usable content width of a <c>w:tc</c> in twips, or null when it can't be estimated.</summary>
    internal static int? EstimateCellContentWidthInDxa(WordDocument document, OfficeXmlNode cell) =>
        WordTable.EstimateCellContentWidthInDxa(document, (W.TableCell)cell.SdkElement!);

    /// <summary>
    /// A cell-width estimator for one <c>w:tbl</c>: given a <c>w:tc</c>, its grid column and its span, the usable content
    /// width in twips.
    /// </summary>
    internal static Func<OfficeXmlNode, int, int, int?> CreateCellContentWidthEstimatorInDxa(WordDocument document, OfficeXmlNode table) {
        Func<W.TableCell, int, int, int?> estimate = WordTable.CreateCellContentWidthEstimatorInDxa(document, (W.Table)table.SdkElement!);
        return (cell, gridIndex, span) => estimate((W.TableCell)cell.SdkElement!, gridIndex, span);
    }

    /// <summary>Adds an internal hyperlink to <paramref name="anchor"/> holding <paramref name="inlineContent"/>.</summary>
    internal static WordParagraph AddHyperLinkContent(WordParagraph paragraph, IEnumerable<OfficeXmlNode> inlineContent, string anchor,
        bool addStyle = false, string tooltip = "", bool history = true, bool moveContent = false) =>
        WordHyperLink.AddHyperLinkContent(paragraph, inlineContent.Select(node => (OpenXmlElement)node.SdkElement!), anchor,
            addStyle, tooltip, history, moveContent);

    /// <summary>Adds an external hyperlink to <paramref name="uri"/> holding <paramref name="inlineContent"/>.</summary>
    internal static WordParagraph AddHyperLinkContent(WordParagraph paragraph, IEnumerable<OfficeXmlNode> inlineContent, Uri uri,
        bool addStyle = false, string tooltip = "", bool history = true, bool moveContent = false) =>
        WordHyperLink.AddHyperLinkContent(paragraph, inlineContent.Select(node => (OpenXmlElement)node.SdkElement!), uri,
            addStyle, tooltip, history, moveContent);

    /// <summary>The story root (document, header, footer, notes or comments) that holds the paragraph.</summary>
    internal static OfficeXmlNode LocationNode(this WordParagraph paragraph) => OfficeXmlNode.FromSdk(paragraph.Location());

    /// <summary>The <c>wp:positionH</c> of an anchored image; throws for an inline image, as the model does.</summary>
    internal static OfficeXmlNode HorizontalPositionNode(this WordImage image) => OfficeXmlNode.FromSdk(image.horizontalPosition);

    // ---- Word.Markdown ---

    /// <summary>The run a run view shows when a field narrowed it to its visible children, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode VisibleRunNode(this WordParagraph run) => OfficeXmlNode.FromSdk(run._visibleRun);

    /// <summary>
    /// The identifier of the first default style of the <c>w:type</c> token <paramref name="type"/>, as
    /// <see cref="WordStyleCatalog.DefaultStyleId"/> reads it.
    /// </summary>
    internal static string? DefaultStyleId(this WordStyleCatalog catalog, string type) => catalog.DefaultStyleId(StyleType(type));

    /// <summary>
    /// A detached <c>w:style</c> holding the theme's table body or header text defaults, as
    /// <see cref="WordThemeDefaults.CreateTableTextStyle"/> builds it.
    /// </summary>
    internal static OfficeXmlNode CreateTableTextStyleNode(OfficeTheme theme, bool header) =>
        OfficeXmlNode.FromSdk(WordThemeDefaults.CreateTableTextStyle(theme, header));

    /// <summary>Whether the <c>w:tc</c> takes the theme's header text, as <see cref="WordDocument.IsThemeHeaderCell"/>.</summary>
    internal static bool IsThemeHeaderCell(OfficeXmlNode cell) => WordDocument.IsThemeHeaderCell((W.TableCell)cell.SdkElement!);

    /// <summary>
    /// Replaces the children of <paramref name="target"/> named as <paramref name="value"/> with a copy of it, in schema
    /// order, as <see cref="WordThemeDefaults.SetVisualChild"/>.
    /// </summary>
    internal static void SetVisualChild(OfficeXmlNode target, OfficeXmlNode value) =>
        WordThemeDefaults.SetVisualChild((OpenXmlCompositeElement)target.SdkElement!, (OpenXmlElement)value.SdkElement!);
}
