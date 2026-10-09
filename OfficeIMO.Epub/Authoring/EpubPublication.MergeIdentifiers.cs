using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static Dictionary<string, string> PrepareMergeIdentifierMap(XElement second, HashSet<string> firstIds,
        HashSet<string> secondIds, HashSet<string> shared, string boundaryId, IReadOnlyDictionary<string, string> supplied,
        CancellationToken token) {
        if (supplied == null || supplied.Count > 10000) throw new ArgumentException("Supply at most 10,000 second-chapter identifier replacements.", nameof(supplied));
        var map = new Dictionary<string, string>(StringComparer.Ordinal);
        var bodyIds = EpubContentIdentifiers.Collect(second.Element(Html + "body")!, "second chapter", true, token);
        foreach (var pair in supplied) {
            token.ThrowIfCancellationRequested();
            if (string.IsNullOrEmpty(pair.Key) || pair.Key.Length > 1024 || string.IsNullOrEmpty(pair.Value) || pair.Value.Length > 1024)
                throw new ArgumentException("Identifier replacements require nonempty names of at most 1024 characters.", nameof(supplied));
            XmlConvert.VerifyNCName(pair.Value);
            if (!bodyIds.Contains(pair.Key) || shared.Contains(pair.Key))
                throw new NotSupportedException("Only second-chapter body identifiers outside shared merge containers can be replaced: " + pair.Key);
            map.Add(pair.Key, pair.Value);
        }
        var final = new HashSet<string>(firstIds, StringComparer.Ordinal) { boundaryId };
        foreach (string id in secondIds.Except(shared, StringComparer.Ordinal)) {
            token.ThrowIfCancellationRequested();
            string replacement = map.TryGetValue(id, out string? value) ? value : id;
            // Equivalent head IDs are retained only once by normal head reconciliation.
            if (bodyIds.Contains(id) && !final.Add(replacement))
                throw new InvalidDataException("The merged body still has a conflicting identifier: " + replacement);
        }
        return map;
    }

    private void VerifyMergeStylesheetFragments(IReadOnlyDictionary<string, string> map, CancellationToken token) {
        if (map.Count == 0) return;
        foreach (EpubManifestItem item in Manifest.Where(item => HasMediaType(item.MediaType, "text/css"))) {
            token.ThrowIfCancellationRequested();
            string path = RequireLocalPath(item);
            if (_encryption.Any(entry => entry.Path == path) ||
                !OfficeIMO.Html.HtmlResourcePipeline.TryDecodeStylesheet(GetResourceBytes(item.Id), "text/css", out string css))
                throw new NotSupportedException("Identifier reconciliation requires inspectable stylesheets: " + path);
            RewriteMovedCss(css, value => {
                token.ThrowIfCancellationRequested();
                if (value.TrimStart().StartsWith("#", StringComparison.Ordinal)) {
                    EpubReference reference = EpubReference.Resolve(path, value);
                    if (reference.Fragment != null && map.TryGetValue(reference.Fragment, out string? replacement) && replacement != reference.Fragment)
                        throw new NotSupportedException("Reconcile the fragment-only stylesheet URL before renaming its chapter identifier: " + value + " in " + path);
                }
                return value;
            }, includeFragmentReferences: true);
        }
    }

    private static void ApplyMergeIdentifierMap(XElement root, IReadOnlyDictionary<string, string> map, CancellationToken token) {
        if (map.Count == 0) return;
        var mapNames = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (XElement element in root.DescendantsAndSelf()) {
            token.ThrowIfCancellationRequested();
            string[] originalIds = element.Attributes().Where(a => a.Name == "id" || a.Name == XNamespace.Xml + "id").Select(a => a.Value).ToArray();
            if ((element.Name == Html + "map" || element.Name == Html + "a") && element.Attribute("name") is XAttribute name &&
                originalIds.Contains(name.Value, StringComparer.Ordinal) && map.TryGetValue(name.Value, out string? namedReplacement)) {
                if (element.Name == Html + "map") mapNames.Add(name.Value, namedReplacement);
                name.Value = namedReplacement;
            }
            foreach (XAttribute attribute in element.Attributes().Where(a => a.Name == "id" || a.Name == XNamespace.Xml + "id"))
                if (map.TryGetValue(attribute.Value, out string? replacement)) attribute.Value = replacement;
        }
        foreach (XElement element in root.DescendantsAndSelf().Where(e => e.Name == Html + "img" || e.Name == Html + "object")) {
            token.ThrowIfCancellationRequested();
            if (element.Attribute("usemap") is XAttribute usemap && usemap.Value.StartsWith("#", StringComparison.Ordinal) &&
                mapNames.TryGetValue(usemap.Value.Substring(1), out string? replacement)) usemap.Value = "#" + replacement;
        }
        EpubContentIdentifiers.RewriteReferences(root, map, token);
    }
}
