using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static void VerifyMergeDocument(XDocument document) {
        XElement root = document.Root!;
        if (root.Elements(Html + "head").Count() != 1 || root.Elements(Html + "body").Count() != 1 ||
            root.Elements().Any(element => element.Name != Html + "head" && element.Name != Html + "body") ||
            root.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value)))
            throw new InvalidDataException("Merging requires one head and one body, without additional root-level content.");
    }

    private static IEnumerable<string> MergeRootNotes(XElement root) => root.Nodes().Where(node => !(node is XElement) &&
        !(node is XText text && string.IsNullOrWhiteSpace(text.Value))).Select(node => node.ToString());

    private static bool SameMergeAttributes(XElement first, XElement second) =>
        first.Attributes().OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal)
            .Select(attribute => new KeyValuePair<XName, string>(attribute.Name, attribute.Value)).SequenceEqual(
                second.Attributes().OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal)
                    .Select(attribute => new KeyValuePair<XName, string>(attribute.Name, attribute.Value)), EqualityComparer<KeyValuePair<XName, string>>.Default);

    private static XElement ComparableMergeHead(XElement head) {
        var copy = new XElement(head);
        foreach (XElement title in copy.Elements(Html + "title")) {
            if (title.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration))
                throw new NotSupportedException("Merging requires unrefined chapter titles without content identifiers or additional attributes.");
            title.Value = string.Empty;
        }
        copy.Nodes().OfType<XText>().Where(text => string.IsNullOrWhiteSpace(text.Value)).Remove();
        return copy;
    }

    private static HashSet<string> FindMergeSeam(XElement first, XElement second, CancellationToken token) {
        var shared = new HashSet<string>(StringComparer.Ordinal);
        for (int depth = 0; ; depth++) {
            token.ThrowIfCancellationRequested();
            if (depth > 64) throw new NotSupportedException("Chapter merge nesting exceeds 64 structural levels.");
            foreach (XAttribute id in second.Attributes().Where(attribute => attribute.Name == "id" || attribute.Name == XNamespace.Xml + "id")) shared.Add(id.Value);
            if (!(first.LastNode is XElement left) || !(second.FirstNode is XElement right) || !CanJoinMergeContainers(left, right)) break;
            first = left; second = right;
        }
        return shared;
    }

    private static bool CanJoinMergeContainers(XElement first, XElement second) => first.Name == second.Name && IsSplitContainer(first) &&
        (first.Attribute("id") != null || first.Attribute(XNamespace.Xml + "id") != null) && SameMergeAttributes(first, second);

    private static void JoinMergeContainers(XElement first, XElement second, XElement boundary, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (first.LastNode is XElement left && second.FirstNode is XElement right && CanJoinMergeContainers(left, right)) {
            JoinMergeContainers(left, right, boundary, token);
            foreach (XNode node in second.Nodes().Skip(1)) first.Add(node);
        } else {
            first.Add(boundary);
            first.Add(second.Nodes());
        }
    }

    private static void VerifyMergeLocalReferences(XElement second, HashSet<string> shared, string path, CancellationToken token) {
        // A shared structural ID survives on the first container; a document-local relationship
        // cannot be redirected to an empty chapter-boundary marker without changing its meaning.
        var local = EpubContentIdentifiers.Collect(second, path, true, token);
        local.ExceptWith(shared);
        EpubContentIdentifiers.ValidateReferences(second, local, path, token);
    }
}
