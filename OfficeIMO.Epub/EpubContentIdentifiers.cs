using System.Threading;

namespace OfficeIMO.Epub;

/// <summary>Document-local identifiers shared by manuscript splitting and authored-content validation.</summary>
internal static class EpubContentIdentifiers {
    internal static HashSet<string> Collect(XElement root, string path, bool rejectDuplicates, CancellationToken token) {
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (XElement element in root.DescendantsAndSelf()) {
            token.ThrowIfCancellationRequested();
            // An element may carry matching id and xml:id attributes; they identify the same target.
            foreach (string id in element.Attributes().Where(attribute => attribute.Name == "id" ||
                attribute.Name == XNamespace.Xml + "id").Select(attribute => attribute.Value).Distinct(StringComparer.Ordinal)) {
                if (!ids.Add(id) && rejectDuplicates)
                    throw new InvalidDataException("Duplicate content id in " + path + ": " + id);
            }
        }
        return ids;
    }

    internal static void ValidateReferences(XElement root, HashSet<string> ids, string path, CancellationToken token) {
        var mapNames = new HashSet<string>(root.DescendantsAndSelf().Where(element =>
            element.Name == XName.Get("map", "http://www.w3.org/1999/xhtml")).Attributes("name").Select(attribute => attribute.Value), StringComparer.Ordinal);
        foreach (XElement element in root.DescendantsAndSelf()) {
            token.ThrowIfCancellationRequested();
            string? map = GetImageMapReference(element);
            if (map != null && !mapNames.Contains(map))
                throw new InvalidDataException("Document-local image map missing in " + path + ": #" + map + ". Keep the image and its map in one chapter.");
            foreach (XAttribute attribute in GetReferenceAttributes(element)) {
                foreach (string id in GetReferenceIds(attribute)) {
                    token.ThrowIfCancellationRequested();
                    if (!ids.Contains(id)) throw new InvalidDataException("Document-local " + attribute.Name.LocalName +
                        " target missing in " + path + ": " + id + ". Keep the related content in one chapter or repair the reference.");
                }
            }
        }
    }

    internal static void RewriteReferences(XElement root, IReadOnlyDictionary<string, string> replacements, CancellationToken token) {
        foreach (XElement element in root.DescendantsAndSelf()) {
            token.ThrowIfCancellationRequested();
            foreach (XAttribute attribute in GetReferenceAttributes(element)) {
                attribute.Value = System.Text.RegularExpressions.Regex.Replace(attribute.Value, @"[^ \t\r\n\f]+", match => {
                    token.ThrowIfCancellationRequested();
                    return replacements.TryGetValue(match.Value, out string? value) ? value : match.Value;
                });
            }
        }
    }

    /// <summary>Enumerates the same document-local relationships used by validation, repair and boundary planning.</summary>
    internal static IEnumerable<XAttribute> GetReferenceAttributes(XElement element) {
        foreach (XAttribute attribute in element.Attributes())
            if (attribute.Name.NamespaceName.Length == 0 && OfficeIMO.Html.HtmlIdentifierReferences.IsReference(
                element.Name.NamespaceName, element.Name.LocalName, attribute.Name.LocalName)) yield return attribute;
    }

    internal static IEnumerable<string> GetReferenceIds(XAttribute attribute) =>
        attribute.Value.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries);

    internal static string? GetImageMapReference(XElement element) =>
        element.Name.NamespaceName == "http://www.w3.org/1999/xhtml" && element.Name.LocalName is "img" or "object" &&
        (string?)element.Attribute("usemap") is string map && map.StartsWith("#", StringComparison.Ordinal) ? map.Substring(1) : null;

}
