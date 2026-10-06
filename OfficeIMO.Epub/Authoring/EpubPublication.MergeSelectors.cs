using OfficeIMO.Html;
using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private sealed class MergeSelectorStyles {
        internal readonly List<XElement> Declarations = new List<XElement>();
        internal readonly Dictionary<string, byte[]> Entries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
    }

    // Work on detached XML and staged resources only. No live package state changes before final validation.
    private MergeSelectorStyles PrepareMergeSelectors(XDocument second, string owner, IReadOnlyDictionary<string, string> ids,
        ContentReferenceMap map, List<(XAttribute Attribute, string Original)> relationships, CancellationToken token) {
        var result = new MergeSelectorStyles();
        if (!ids.Any(pair => pair.Key != pair.Value)) return result;
        if (second.DescendantNodes().OfType<XProcessingInstruction>().Any())
            throw new NotSupportedException("Selector reconciliation requires stylesheet elements instead of processing instructions.");
        var relationshipSelectors = new MergeRelationshipSelectors(relationships, token);
        var clones = new Dictionary<string, string>(StringComparer.Ordinal);
        var usedIds = new HashSet<string>(Root.DescendantsAndSelf().Attributes("id").Select(attribute => attribute.Value), StringComparer.Ordinal);
        var usedPaths = new HashSet<string>(_entries.Keys.Concat(new[] { PackagePath })
            .Select(path => path.Normalize(NormalizationForm.FormC)), StringComparer.OrdinalIgnoreCase);
        int number = 0;
        long clonedBytes = 0;
        foreach (XElement element in second.Descendants()) {
            token.ThrowIfCancellationRequested();
            if (element.Name.LocalName == "style" && (element.Name.Namespace == Html || element.Name.NamespaceName == "http://www.w3.org/2000/svg")) {
                if (element.Attribute("type") is XAttribute type && !HasMediaType(type.Value, "text/css"))
                    throw new NotSupportedException("Selector reconciliation requires CSS style elements.");
                element.Value = Rewrite(element.Value, owner, 0);
            } else if (element.Name == Html + "link" && Tokens((string?)element.Attribute("rel"))
                .Any(value => value.Equals("stylesheet", StringComparison.OrdinalIgnoreCase))) {
                XAttribute href = element.Attribute("href") ?? throw new InvalidDataException("Stylesheet link has no href.");
                href.Value = Import(href.Value, owner, 0);
                // Cloned bytes no longer have the original integrity digest.
                if (element.Attribute("integrity") != null) throw new NotSupportedException("Reconcile stylesheet integrity metadata before cloning styles.");
            }
        }
        return result;

        string Rewrite(string css, string path, int depth) {
            string selectors = HtmlCssIdSelectorRewriter.Rewrite(css, ids, token, relationshipSelectors.Rewrite);
            return HtmlResourcePipeline.RewriteCssResourceUrls(selectors, (value, kind) => {
                token.ThrowIfCancellationRequested();
                return kind == HtmlResourceKind.Stylesheet ? Import(value, path, depth) :
                    RewriteMovedReference(path, null, path, null, value, string.Empty, string.Empty, map);
            }, includeFragmentReferences: true);
        }

        string Import(string value, string path, int depth) {
            token.ThrowIfCancellationRequested();
            if (depth >= 64) throw new NotSupportedException("Selector reconciliation exceeds 64 stylesheet import levels.");
            EpubReference reference = EpubReference.Resolve(path, value);
            if (reference.Kind != EpubReferenceKind.Container || reference.ContainerPath == null)
                throw new NotSupportedException("Selector reconciliation requires local declared stylesheets.");
            string source = reference.ContainerPath;
            if (!clones.TryGetValue(source, out string? destination)) {
                EpubManifestItem[] declarations = Manifest.Where(item => item.Reference.ContainerPath == source).ToArray();
                if (declarations.Length != 1 || !HasMediaType(declarations[0].MediaType, "text/css"))
                    throw new NotSupportedException("Selector reconciliation requires one CSS declaration per stylesheet.");
                EpubManifestItem item = declarations[0];
                XElement declaration = RequireSection("manifest").Elements(Opf + "item").Single(entry => (string?)entry.Attribute("id") == item.Id);
                if (declaration.Elements().Any() || declaration.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                    attribute.Name != "id" && attribute.Name != "href" && attribute.Name != "media-type") ||
                    Root.Descendants().Attributes("refines").Any(attribute => ReferencesPackageId(attribute.Value, item.Id)))
                    throw new NotSupportedException("Stylesheet declaration metadata requires explicit reconciliation before cloning.");
                EnsureResourceMutationAllowed(source);
                if (!HtmlResourcePipeline.TryDecodeStylesheet(_entries[source], "text/css", out string css))
                    throw new InvalidDataException("Stylesheet cannot be decoded: " + source);
                string directory = source.Substring(0, source.LastIndexOf('/') + 1);
                string id;
                do { number++; id = "merge-style-" + number.ToString(System.Globalization.CultureInfo.InvariantCulture); destination = directory + id + ".css"; }
                while (usedIds.Contains(id) || usedPaths.Contains(destination.Normalize(NormalizationForm.FormC)));
                usedIds.Add(id); usedPaths.Add(destination.Normalize(NormalizationForm.FormC)); clones.Add(source, destination);
                EnsureEntryBudget(clones.Count - 1); // The second chapter is removed in the same transaction.
                result.Declarations.Add(new XElement(Opf + "item", new XAttribute("id", id),
                    new XAttribute("href", RelativeHref(PackagePath, destination)), new XAttribute("media-type", "text/css")));
                // Clone in the source directory, so ordinary relative resource URLs retain their base.
                string rewritten = Rewrite(css, source, depth + 1);
                if (rewritten.TrimStart().StartsWith("@charset", StringComparison.OrdinalIgnoreCase)) {
                    int end = rewritten.IndexOf(';');
                    if (end >= 0) rewritten = "@charset \"UTF-8\";" + rewritten.Substring(end + 1);
                }
                byte[] bytes = new UTF8Encoding(false, true).GetBytes(rewritten);
                if (bytes.LongLength > _maximumEntryBytes) throw new InvalidDataException("Cloned stylesheet exceeds the entry-byte limit.");
                result.Entries.Add(destination, bytes);
                clonedBytes += bytes.LongLength;
                if (clonedBytes > _maximumRetainedBytes)
                    throw new InvalidDataException("Cloned stylesheets exceed the retained-byte limit.");
            }
            return RelativeHref(path, destination) + (reference.Query == null ? string.Empty : "?" + reference.Query) +
                (reference.Fragment == null ? string.Empty : "#" + Uri.EscapeDataString(reference.Fragment));
        }
    }
}
