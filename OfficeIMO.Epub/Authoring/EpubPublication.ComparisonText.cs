using System.Threading;

namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private static void CompareEditionText(string manifestId, XDocument before, XDocument after,
        List<EpubEditionTextChange> output, CancellationToken token) {
        var left = EditionTextBlocks(before, token); var right = EditionTextBlocks(after, token);
        foreach (string locator in left.Keys.Union(right.Keys, StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal)) {
            token.ThrowIfCancellationRequested();
            left.TryGetValue(locator, out string? previous); right.TryGetValue(locator, out string? current);
            if (previous == current) continue;
            if (output.Count >= 10_000) throw new InvalidDataException("Edition comparison exceeds the 10,000 changed text-block bound.");
            output.Add(new EpubEditionTextChange(manifestId, locator, previous, current));
        }
    }

    private static Dictionary<string, string> EditionTextBlocks(XDocument document, CancellationToken token) {
        XElement? body = document.Root?.Element(Html + "body");
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        if (body == null) return result;
        var names = new HashSet<string>(new[] { "p", "h1", "h2", "h3", "h4", "h5", "h6", "li", "dt", "dd", "blockquote", "pre", "td", "th", "figcaption" }, StringComparer.Ordinal);
        var all = body.Descendants().Where(element => element.Name.Namespace == Html && names.Contains(element.Name.LocalName)).ToArray();
        var blocks = all.Where(element => !element.Descendants().Any(child => child.Name.Namespace == Html && names.Contains(child.Name.LocalName))).ToArray();
        var idCounts = body.DescendantsAndSelf().Select(element => (string?)element.Attribute("id") ?? (string?)element.Attribute(XNamespace.Xml + "id"))
            .Where(id => !string.IsNullOrEmpty(id)).GroupBy(id => id!, StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        Dictionary<XElement, int>? siblingPositions = null;
        foreach (XElement element in blocks) {
            token.ThrowIfCancellationRequested();
            string? id = (string?)element.Attribute("id") ?? (string?)element.Attribute(XNamespace.Xml + "id");
            string locator;
            if (!string.IsNullOrEmpty(id) && idCounts[id!] == 1) {
                locator = "#" + id;
            } else {
                siblingPositions ??= EditionSiblingPositions(document.Root!, token);
                locator = "/" + string.Join("/", element.AncestorsAndSelf().Reverse()
                    .Select(node => node.Name + "[" + siblingPositions[node] + "]"));
            }
            result.Add(locator, element.Value);
        }
        // Retain a whole-body signal as well, covering loose text outside selected blocks.
        result.Add("$body", body.Value);
        return result;
    }

    private static Dictionary<XElement, int> EditionSiblingPositions(XElement root, CancellationToken token) {
        var positions = new Dictionary<XElement, int> { [root] = 1 };
        foreach (XElement parent in root.DescendantsAndSelf()) {
            token.ThrowIfCancellationRequested();
            var counts = new Dictionary<XName, int>();
            foreach (XElement child in parent.Elements()) {
                counts.TryGetValue(child.Name, out int count);
                positions.Add(child, count + 1);
                counts[child.Name] = count + 1;
            }
        }
        return positions;
    }
}
