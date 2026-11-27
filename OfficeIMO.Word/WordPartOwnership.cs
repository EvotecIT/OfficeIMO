using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.OpenXml.Internal;

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

    // Part plumbing for SmartArt, whose markup is read and written without SDK types. The part handles are opaque
    // here until Word parts are held by OfficeIMO's own package layer.

    /// <summary>The story part that owns <paramref name="element"/>, as an opaque part handle.</summary>
    internal static object ResolveOwner(WordDocument document, OfficeXmlNode element) =>
        Resolve(document, (OpenXmlElement)(element.SdkElement ?? throw new InvalidOperationException("The element is not held by the SDK package.")));

    /// <summary>Adds a new SmartArt part under <paramref name="owner"/> with <paramref name="content"/> and returns its relationship id.</summary>
    internal static string AddDiagramPart(object owner, WordDiagramPartKind kind, byte[] content) {
        var part = (OpenXmlPart)owner;
        OpenXmlPart added = kind switch {
            WordDiagramPartKind.Layout => part.AddNewPart<DiagramLayoutDefinitionPart>(),
            WordDiagramPartKind.Colors => part.AddNewPart<DiagramColorsPart>(),
            WordDiagramPartKind.Style => part.AddNewPart<DiagramStylePart>(),
            WordDiagramPartKind.Data => part.AddNewPart<DiagramDataPart>(),
            _ => part.AddNewPart<DiagramPersistLayoutPart>()
        };
        using (var stream = new MemoryStream(content, writable: false)) added.FeedData(stream);
        return part.GetIdOfPart(added);
    }

    /// <summary>The SmartArt data part <paramref name="relationshipId"/> names under <paramref name="owner"/>, or null.</summary>
    internal static object? GetDiagramDataPart(object owner, string relationshipId) =>
        ((OpenXmlPart)owner).GetPartById(relationshipId) as DiagramDataPart;

    /// <summary>Reads a SmartArt data part: its stream, or the loaded data model when the stream is empty.</summary>
    internal static XDocument ReadDiagramData(object dataPart) {
        var part = (DiagramDataPart)dataPart;
        using (var stream = part.GetStream(FileMode.Open, FileAccess.Read)) {
            if (stream != null && stream.Length > 0) return XDocument.Load(stream);
        }
        if (part.DataModelRoot != null) return XDocument.Parse(part.DataModelRoot.OuterXml);
        throw new InvalidOperationException("SmartArt data part is empty.");
    }

    /// <summary>Replaces the content of a SmartArt data part.</summary>
    internal static void WriteDiagramData(object dataPart, XDocument data) {
        using var stream = new MemoryStream();
        data.Save(stream);
        stream.Position = 0;
        ((DiagramDataPart)dataPart).FeedData(stream);
    }

    /// <summary>The unique id of the SmartArt layout definition <paramref name="relationshipId"/> names, or null.</summary>
    internal static string? GetDiagramLayoutUniqueId(object owner, string relationshipId) {
        if (((OpenXmlPart)owner).GetPartById(relationshipId) is not DiagramLayoutDefinitionPart part || part.LayoutDefinition == null) return null;
        return part.LayoutDefinition.UniqueId?.Value ?? string.Empty;
    }

    /// <summary>The root of the SmartArt persisted layout part <paramref name="relationshipId"/> names, or none.</summary>
    internal static OfficeXmlNode GetDiagramPersistedLayout(object owner, string relationshipId) =>
        ((OpenXmlPart)owner).GetPartById(relationshipId) is DiagramPersistLayoutPart part ? OfficeXmlNode.FromSdk(part.Drawing) : default;
}

/// <summary>The parts a SmartArt diagram keeps next to the story that shows it.</summary>
internal enum WordDiagramPartKind {
    Layout,
    Colors,
    Style,
    Data,
    PersistLayout
}
