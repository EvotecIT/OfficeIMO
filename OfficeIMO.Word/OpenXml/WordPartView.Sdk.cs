using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.OpenXml;

/// <summary>
/// The SDK side of <see cref="WordPackageView"/>: the one place where readers that walk a Word package's markup meet a
/// package the Open XML SDK holds. Everything here takes and returns <see cref="object"/> or OfficeIMO types, so the
/// readers name no SDK type. It goes away with the SDK.
/// </summary>
internal sealed partial class WordPackageView {
    private static object SdkPackageOf(WordDocument document) => document._wordprocessingDocument;

    private static object? SdkMainPart(object document) => ((WordprocessingDocument)document).MainDocumentPart;

    internal static string SdkPartName(object part) => ((OpenXmlPart)part).Uri.ToString();

    internal static string? SdkContentType(object part) => ((OpenXmlPart)part).ContentType;

    internal static object? SdkRootElement(object part) => ((OpenXmlPart)part).RootElement;

    /// <summary>Whether the part is a styles part the SDK has not loaded, in a package that reads markup as written.</summary>
    internal static bool SdkIsUnloadedStylesPart(object part) =>
        part is StylesPart stylesPart && !stylesPart.IsRootElementLoaded &&
        stylesPart.OpenXmlPackage.MarkupCompatibilityProcessSettings.ProcessMode == DocumentFormat.OpenXml.Packaging.MarkupCompatibilityProcessMode.NoProcess;

    internal static List<WordPartRelationship> SdkRelationships(object part) {
        var owner = (OpenXmlPart)part;
        var relationships = new List<WordPartRelationship>();
        foreach (IdPartPair pair in owner.Parts) {
            relationships.Add(new WordPartRelationship(pair.RelationshipId, pair.OpenXmlPart.RelationshipType, pair.OpenXmlPart.Uri.ToString(), isExternal: false));
        }
        foreach (ExternalRelationship relationship in owner.ExternalRelationships) {
            relationships.Add(new WordPartRelationship(relationship.Id, relationship.RelationshipType, relationship.Uri.OriginalString, isExternal: true));
        }
        foreach (HyperlinkRelationship relationship in owner.HyperlinkRelationships) {
            relationships.Add(new WordPartRelationship(relationship.Id, HyperlinkRelationshipType, relationship.Uri.OriginalString, relationship.IsExternal));
        }
        return relationships;
    }

    internal static object? SdkPartById(object part, string id) {
        try {
            return ((OpenXmlPart)part).GetPartById(id);
        } catch (ArgumentOutOfRangeException) {
            return null;
        }
    }

    internal static byte[] SdkReadBytes(object part) {
        using Stream stream = ((OpenXmlPart)part).GetStream(FileMode.Open, FileAccess.Read);
        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    internal static string SdkAddImagePart(object part, string contentType, byte[] content) {
        var owner = (OpenXmlPart)part;
        ImagePart imagePart = owner switch {
            MainDocumentPart mainDocumentPart => mainDocumentPart.AddImagePart(contentType),
            HeaderPart headerPart => headerPart.AddImagePart(contentType),
            FooterPart footerPart => footerPart.AddImagePart(contentType),
            FootnotesPart footnotesPart => footnotesPart.AddImagePart(contentType),
            EndnotesPart endnotesPart => endnotesPart.AddImagePart(contentType),
            _ => throw new InvalidOperationException("Images cannot be copied into this document part.")
        };
        using (var stream = new MemoryStream(content, writable: false)) imagePart.FeedData(stream);
        return owner.GetIdOfPart(imagePart);
    }

    internal static string SdkAddExternalRelationship(object part, string type, string uri) =>
        ((OpenXmlPart)part).AddExternalRelationship(type, new Uri(uri, UriKind.RelativeOrAbsolute)).Id;

    internal static string SdkAddHyperlinkRelationship(object part, string uri, bool isExternal) =>
        ((OpenXmlPart)part).AddHyperlinkRelationship(new Uri(uri, UriKind.RelativeOrAbsolute), isExternal).Id;

    // ---- Elements -------------------------------------------------------------------------------------------

    /// <summary>A date and time as the SDK writes a <c>w:date</c> value.</summary>
    internal static string FormatDateTime(DateTime value) => new DocumentFormat.OpenXml.DateTimeValue(value).InnerText ?? string.Empty;

    /// <summary>The attributes of an SDK element as the SDK reports them: without namespace declarations.</summary>
    internal static void SdkAttributes(object element, List<WordXmlAttribute> attributes) {
        foreach (DocumentFormat.OpenXml.OpenXmlAttribute attribute in ((DocumentFormat.OpenXml.OpenXmlElement)element).GetAttributes()) {
            attributes.Add(new WordXmlAttribute(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, attribute.Value ?? string.Empty));
        }
    }

    internal static void SdkSetAttribute(object element, string prefix, string localName, string namespaceUri, string value) =>
        ((DocumentFormat.OpenXml.OpenXmlElement)element).SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute(prefix, localName, namespaceUri, value));

    internal static void SdkRemoveAttribute(object element, string localName, string namespaceUri) =>
        ((DocumentFormat.OpenXml.OpenXmlElement)element).RemoveAttribute(localName, namespaceUri);

    internal static string? SdkLookupNamespace(object element, string prefix) => ((DocumentFormat.OpenXml.OpenXmlElement)element).LookupNamespace(prefix);

    private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<WordPartView, System.Runtime.CompilerServices.StrongBox<Numbering?>> NumberingByPart = new();

    /// <summary>
    /// The numbering definition text of a list item as reports carry it: copies of its <c>w:num</c> and <c>w:abstractNum</c>
    /// with volatile and, unless they are compared, generated identifiers cleared, then the level and level override in
    /// use, each written as the SDK writes markup it read from the numbering part. A part of OfficeIMO's own package is read
    /// into an SDK numbering element once, as the SDK reads it from a package, since reading markup out of a package
    /// context spells some attributes differently.
    /// </summary>
    internal static string SdkListNumberingSignature(WordPartView? numberingPart, string numberId, string level, bool compareGeneratedIds) {
        Numbering? numbering = GetSdkNumbering(numberingPart);
        if (numbering == null || string.IsNullOrWhiteSpace(numberId)) {
            return string.Empty;
        }

        NumberingInstance? instance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(item => string.Equals(
                item.NumberID?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                numberId,
                StringComparison.Ordinal));
        if (instance == null) {
            return string.Empty;
        }

        int? abstractNumberId = instance.AbstractNumId?.Val?.Value;
        AbstractNum? abstractNum = abstractNumberId.HasValue
            ? numbering.Elements<AbstractNum>().FirstOrDefault(item => item.AbstractNumberId?.Value == abstractNumberId.Value)
            : null;

        var instanceClone = (NumberingInstance)instance.CloneNode(true);
        var abstractClone = (AbstractNum?)abstractNum?.CloneNode(true);
        if (abstractClone != null) {
            abstractClone.RemoveAllChildren<Nsid>();
            abstractClone.RemoveAllChildren<TemplateCode>();
            foreach (Level item in abstractClone.Descendants<Level>()) {
                item.TemplateCode = null;
            }
        }

        if (!compareGeneratedIds) {
            instanceClone.NumberID = 0;
            AbstractNumId? abstractNumId = instanceClone.GetFirstChild<AbstractNumId>();
            if (abstractNumId != null) {
                abstractNumId.Val = 0;
            }

            if (abstractClone != null) {
                abstractClone.AbstractNumberId = 0;
            }
        }

        Level? resolvedLevel = abstractClone?
            .Elements<Level>()
            .FirstOrDefault(item => string.Equals(
                item.LevelIndex?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                level,
                StringComparison.Ordinal));
        LevelOverride? resolvedOverride = instanceClone
            .Elements<LevelOverride>()
            .FirstOrDefault(item => string.Equals(
                item.LevelIndex?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture),
                level,
                StringComparison.Ordinal));

        return string.Join(
            "|",
            instanceClone.OuterXml,
            abstractClone?.OuterXml ?? string.Empty,
            resolvedLevel?.OuterXml ?? string.Empty,
            resolvedOverride?.OuterXml ?? string.Empty);
    }

    /// <summary>
    /// The numbering definition a paragraph's numbering refers to, as paragraph match text carries it: the
    /// <c>w:abstractNum</c> and <c>w:num</c> with their identifiers cleared, written as the SDK writes markup it read
    /// from the numbering part, or the identifier that could not be resolved. Similarity scores read this text, so it is
    /// kept as it was.
    /// </summary>
    internal static string SdkNumberingDefinitionSignature(WordPartView? numberingPart, int numberingId, string level) {
        Numbering? numbering = GetSdkNumbering(numberingPart);
        if (numbering == null) {
            return numberingId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        NumberingInstance? instance = numbering.Elements<NumberingInstance>()
            .FirstOrDefault(item => item.NumberID?.Value == numberingId);
        if (instance == null) {
            return numberingId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        int? abstractNumberId = instance.AbstractNumId?.Val?.Value;
        if (abstractNumberId == null) {
            return numberingId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        AbstractNum? abstractNum = numbering.Elements<AbstractNum>()
            .FirstOrDefault(item => item.AbstractNumberId?.Value == abstractNumberId.Value);
        if (abstractNum == null) {
            return abstractNumberId.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        var clonedAbstract = (AbstractNum)abstractNum.CloneNode(true);
        clonedAbstract.AbstractNumberId = 0;
        var clonedInstance = (NumberingInstance)instance.CloneNode(true);
        clonedInstance.NumberID = 0;
        if (clonedInstance.AbstractNumId != null) {
            clonedInstance.AbstractNumId.Val = 0;
        }

        return clonedAbstract.OuterXml + "|instance:" + clonedInstance.OuterXml + "|level:" + level;
    }

    private static Numbering? GetSdkNumbering(WordPartView? numberingPart) =>
        numberingPart == null ? null
        : numberingPart.SdkPart is NumberingDefinitionsPart sdkPart ? sdkPart.Numbering
        : NumberingByPart.GetValue(numberingPart, ReadNumbering).Value;

    // The numbering part's content read through a numbering part of a small in-memory package.
    private static System.Runtime.CompilerServices.StrongBox<Numbering?> ReadNumbering(WordPartView part) {
        using var stream = new MemoryStream();
        using WordprocessingDocument package = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document);
        NumberingDefinitionsPart numberingPart = package.AddMainDocumentPart().AddNewPart<NumberingDefinitionsPart>();
        using (var content = new MemoryStream(part.ReadBytes(), writable: false)) numberingPart.FeedData(content);
        return new System.Runtime.CompilerServices.StrongBox<Numbering?>((Numbering?)numberingPart.Numbering?.CloneNode(true));
    }

    // ---- Engine state the SDK package carries ---------------------------------------------------------------

    /// <summary>Gives drawings under <paramref name="element"/> identifiers unused in the SDK package.</summary>
    internal static void ReassignDrawingIds(object sdkDocument, OfficeXmlNode element) =>
        WordDrawingIdAllocator.Reassign((WordprocessingDocument)sdkDocument, (DocumentFormat.OpenXml.OpenXmlElement)element.SdkElement!);

    /// <summary>Carries table grid normalization state from SDK elements to their copies.</summary>
    internal static void CopyGridNormalizationState(OfficeXmlNode source, OfficeXmlNode copy) {
        if (source.SdkElement is DocumentFormat.OpenXml.OpenXmlElement original && copy.SdkElement is DocumentFormat.OpenXml.OpenXmlElement cloned) {
            WordTable.CopyGridNormalizationState(original, cloned);
        }
    }

    /// <summary>Reads a drawing's layout. Markup outside the SDK is read through a detached SDK drawing.</summary>
    internal static bool TryReadDrawingLayout(OfficeXmlNode drawing, out WordDrawingLayoutInfo? layout) {
        if (drawing.SdkElement is DocumentFormat.OpenXml.Wordprocessing.Drawing sdkDrawing) return WordDrawingLayoutReader.TryRead(sdkDrawing, out layout);
        return WordDrawingLayoutReader.TryRead(new DocumentFormat.OpenXml.Wordprocessing.Drawing(System.Text.Encoding.UTF8.GetString(drawing.OuterXml())), out layout);
    }

    /// <summary>
    /// Copies <paramref name="target"/>'s package into a new document for an in-place redline, as the target stands,
    /// with its table grid normalization state.
    /// </summary>
    internal static WordDocument CloneForRedline(WordDocument target) {
        using var snapshot = new MemoryStream();
        using (WordPackageSnapshot.Clone(target._wordprocessingDocument, snapshot)) { }
        snapshot.Position = 0;
        WordDocument document = WordDocument.LoadPackageSnapshot(snapshot, target._wordprocessingDocument);
        try {
            WordTable.CopyGridNormalizationState(target._wordprocessingDocument, document._wordprocessingDocument);
            return document;
        } catch {
            document.Dispose();
            throw;
        }
    }

    // ---- Elements behind engine objects --------------------------------------------------------------------

    internal static OfficeXmlNode ElementOf(WordParagraph paragraph) => OfficeXmlNode.FromSdk(paragraph._paragraph);

    internal static OfficeXmlNode ElementOf(WordTable table) => OfficeXmlNode.FromSdk(table._table);

    internal static OfficeXmlNode ElementOf(WordTableRow row) => OfficeXmlNode.FromSdk(row._tableRow);

    internal static OfficeXmlNode ElementOf(WordTableCell cell) => OfficeXmlNode.FromSdk(cell._tableCell);

    /// <summary>The body of the SDK package.</summary>
    internal static OfficeXmlNode SdkBody(WordDocument document) => OfficeXmlNode.FromSdk(document._wordprocessingDocument.MainDocumentPart!.Document!.Body!);
}
