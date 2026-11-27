using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.OpenXml.Internal;

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
internal static partial class WordConverterNodes {
    // ---- Document parts -------------------------------------------------------------------------------------------

    /// <summary>The <c>w:document</c> element.</summary>
    internal static OfficeXmlNode DocumentNode(this WordDocument document) => OfficeXmlNode.FromSdk(document._document);

    /// <summary>The <c>w:body</c> element, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode BodyNode(this WordDocument document) => OfficeXmlNode.FromSdk(document._document.Body);

    private static MainDocumentPart? Main(WordDocument? document) => document?._wordprocessingDocument?.MainDocumentPart;

    /// <summary>The <c>w:settings</c> root, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode SettingsNode(this WordDocument? document) =>
        OfficeXmlNode.FromSdk(Main(document)?.DocumentSettingsPart?.Settings);

    /// <summary>The <c>w:styles</c> root, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode StylesNode(this WordDocument? document) =>
        OfficeXmlNode.FromSdk(Main(document)?.StyleDefinitionsPart?.Styles);

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

    /// <summary>
    /// The part that holds <paramref name="element"/>, as an opaque key: equal for elements of the same part, null for
    /// a detached element.
    /// </summary>
    internal static object? PartKey(this OfficeXmlNode element) =>
        element.SdkElement is { } sdk ? OfficeSdkXml.PartOf(sdk) : element.Tree;

    /// <summary>The root element of the part that holds <paramref name="element"/>.</summary>
    internal static OfficeXmlNode PartRoot(this OfficeXmlNode element) =>
        OfficeXmlNode.FromSdk((PartOf(element))?.RootElement);

    /// <summary>The root element of the part <paramref name="relationshipId"/> names from the part that holds <paramref name="element"/>.</summary>
    internal static OfficeXmlNode RelatedPartRoot(this OfficeXmlNode element, string? relationshipId) =>
        OfficeXmlNode.FromSdk(RelatedPart(element, relationshipId)?.RootElement);

    /// <summary>The root element of the part <paramref name="relationshipId"/> names from the main document part.</summary>
    internal static OfficeXmlNode RelatedPartRoot(this WordDocument document, string? relationshipId) =>
        OfficeXmlNode.FromSdk(TryGetPart(Main(document), relationshipId)?.RootElement);

    /// <summary>The bytes of the part <paramref name="relationshipId"/> names from the part that holds <paramref name="element"/>.</summary>
    internal static byte[]? RelatedPartBytes(this OfficeXmlNode element, string? relationshipId) {
        OpenXmlPart? part = RelatedPart(element, relationshipId);
        if (part == null) return null;
        using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    /// <summary>The content type of the part <paramref name="relationshipId"/> names, or null.</summary>
    internal static string? RelatedPartContentType(this OfficeXmlNode element, string? relationshipId) =>
        RelatedPart(element, relationshipId)?.ContentType;

    /// <summary>The URI of the part <paramref name="relationshipId"/> names, or null.</summary>
    internal static string? RelatedPartUri(this OfficeXmlNode element, string? relationshipId) =>
        RelatedPart(element, relationshipId)?.Uri.OriginalString;

    /// <summary>The relationship type of the part <paramref name="relationshipId"/> names, or null.</summary>
    internal static string? RelatedPartRelationshipType(this OfficeXmlNode element, string? relationshipId) =>
        RelatedPart(element, relationshipId)?.RelationshipType;

    /// <summary>
    /// The target of an external relationship (a hyperlink, a linked image) from the part that holds
    /// <paramref name="element"/>, or null.
    /// </summary>
    internal static Uri? ExternalRelationshipTarget(this OfficeXmlNode element, string? relationshipId) {
        if (string.IsNullOrEmpty(relationshipId) || PartOf(element) is not { } part) return null;
        foreach (HyperlinkRelationship relationship in part.HyperlinkRelationships) {
            if (relationship.Id == relationshipId) return relationship.Uri;
        }
        foreach (ExternalRelationship relationship in part.ExternalRelationships) {
            if (relationship.Id == relationshipId) return relationship.Uri;
        }
        return null;
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

    /// <summary>The simple field this paragraph object stands in, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode SimpleFieldNode(this WordParagraph paragraph) => OfficeXmlNode.FromSdk(paragraph._simpleField);

    /// <summary>The bookmark start this paragraph object stands for, or <see cref="OfficeXmlNode.None"/>.</summary>
    internal static OfficeXmlNode BookmarkStartNode(this WordParagraph paragraph) => OfficeXmlNode.FromSdk(paragraph._bookmarkStart);

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
}
