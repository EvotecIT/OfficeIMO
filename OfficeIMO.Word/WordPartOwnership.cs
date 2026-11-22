using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

/// <summary>Resolves relationship ownership from the containing Word story.</summary>
internal static class WordPartOwnership {
    private static readonly System.Threading.AsyncLocal<DetachedOwner?> ActiveDetachedOwner = new();

    private sealed class DetachedOwner {
        internal DetachedOwner(OpenXmlElement root, OpenXmlPart part, DetachedOwner? previous) {
            Root = root;
            Part = part;
            Previous = previous;
        }
        internal OpenXmlElement Root { get; }
        internal OpenXmlPart Part { get; }
        internal DetachedOwner? Previous { get; }
    }

    private sealed class DetachedOwnerScope : IDisposable {
        private readonly DetachedOwner? _previous;
        internal DetachedOwnerScope(DetachedOwner? previous) => _previous = previous;
        public void Dispose() => ActiveDetachedOwner.Value = _previous;
    }

    /// <summary>
    /// Lets a detached copy of story content, such as a projected frame, resolve its
    /// relationships against the part the content was copied from.
    /// </summary>
    internal static IDisposable UseDetachedOwner(OpenXmlElement root, OpenXmlPart part) {
        DetachedOwner? previous = ActiveDetachedOwner.Value;
        ActiveDetachedOwner.Value = new DetachedOwner(root, part, previous);
        return new DetachedOwnerScope(previous);
    }

    internal static OpenXmlPart Resolve(WordDocument document, OpenXmlElement element) {
        OpenXmlElement? parent = element;
        OpenXmlElement top = element;
        while (parent != null && parent is not Body && parent is not Header && parent is not Footer &&
            parent is not Footnotes && parent is not Endnotes && parent is not Comments) {
            top = parent;
            parent = parent.Parent;
        }
        if (parent == null) {
            for (DetachedOwner? owner = ActiveDetachedOwner.Value; owner != null; owner = owner.Previous)
                if (ReferenceEquals(owner.Root, top)) return owner.Part;
        }
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
