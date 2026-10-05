using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static bool IsSplitContainer(XElement element) => element.Name.Namespace == Html &&
        new[] { "body", "div", "section", "article", "main" }.Contains(element.Name.LocalName, StringComparer.Ordinal);

    private static bool HasSplitContent(XElement element) => element.Nodes().Any(node => node is XElement ||
        node is XText text && !string.IsNullOrWhiteSpace(text.Value));

    private static (XElement First, XElement Second) SplitContentContainer(XElement container, XElement boundary,
        HashSet<XElement> ancestors, CancellationToken token) {
        var first = new XElement(container.Name, container.Attributes());
        var second = new XElement(container.Name, container.Attributes());
        bool after = false;
        foreach (XNode node in container.Nodes()) {
            token.ThrowIfCancellationRequested();
            if (ReferenceEquals(node, boundary)) after = true;
            if (!after && node is XElement child && ancestors.Contains(child)) {
                var split = SplitContentContainer(child, boundary, ancestors, token);
                if (HasSplitContent(split.First)) first.Add(split.First);
                else split.Second.AddFirst(split.First.Nodes());
                second.Add(split.Second);
                after = true;
            } else (after ? second : first).Add(node);
        }
        return (first, second);
    }

    private XElement RequireSplittablePosition(EpubManifestItem item) {
        string path = RequireLocalPath(item);
        if (!HasMediaType(item.MediaType, "application/xhtml+xml") || HasToken(item.Properties, "nav"))
            throw new NotSupportedException("Chapter splitting requires a non-navigation XHTML chapter.");
        if (_rootfilePaths.Length != 1 || _encryption.Count != 0 || Manifest.Any(resource => HasToken(resource.Properties, "scripted") || IsScriptMediaType(resource.MediaType)))
            throw new NotSupportedException("Chapter splitting requires one unencrypted, non-scripted rendition.");
        if (Manifest.Count(resource => resource.Reference.ContainerPath == path) != 1 || item.MediaOverlayId != null || item.FallbackId != null ||
            Manifest.Any(resource => resource.FallbackId == item.Id))
            throw new NotSupportedException("Shared resources, fallbacks and media-overlay chapters require an explicit split policy.");
        XElement[] positions = RequireSection("spine").Elements(Opf + "itemref").Where(element => (string?)element.Attribute("idref") == item.Id).ToArray();
        if (positions.Length != 1) throw new InvalidOperationException("A split chapter must have exactly one spine position.");
        const string rendition = "http://www.idpf.org/vocab/rendition/#";
        string[] properties = Tokens((string?)positions[0].Attribute("properties")).Select(token => EpubVocabulary.Expand(Root, token)).ToArray();
        bool fixedLayout = RequireSection("metadata").Elements(Opf + "meta").Any(meta => meta.Attribute("refines") == null &&
            EpubVocabulary.Expand(Root, (string?)meta.Attribute("property") ?? string.Empty) == rendition + "layout" && meta.Value == "pre-paginated");
        if (properties.Contains(rendition + "layout-pre-paginated") || fixedLayout && !properties.Contains(rendition + "layout-reflowable"))
            throw new NotSupportedException("Fixed-layout chapters require separate geometry-aware splitting.");
        return positions[0];
    }
}
