using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>Resolves relationship ownership from the containing Word story.</summary>
internal static class WordPartOwnership {
    internal static OpenXmlPart Resolve(WordDocument document, OpenXmlElement element) {
        OpenXmlElement? parent = element;
        while (parent != null && parent is not Body && parent is not Header && parent is not Footer &&
            parent is not Footnotes && parent is not Endnotes && parent is not Comments)
            parent = parent.Parent;
        if (parent is Header header)
            return header.HeaderPart ?? throw new InvalidOperationException("Header part is missing.");
        if (parent is Footer footer)
            return footer.FooterPart ?? throw new InvalidOperationException("Footer part is missing.");
        var main = document.MainDocumentPartRoot;
        if (parent is Footnotes)
            return main.FootnotesPart ?? throw new InvalidOperationException("FootnotesPart is missing.");
        if (parent is Endnotes)
            return main.EndnotesPart ?? throw new InvalidOperationException("EndnotesPart is missing.");
        if (parent is Comments)
            return main.WordprocessingCommentsPart ?? throw new InvalidOperationException("WordprocessingCommentsPart is missing.");
        return main;
    }
}
