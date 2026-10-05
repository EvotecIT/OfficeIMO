using System.Threading;

namespace OfficeIMO.Epub;

/// <summary>Document-local identifiers shared by manuscript splitting and authored-content validation.</summary>
internal static class EpubContentIdentifiers {
    private static readonly HashSet<string> ReferenceAttributes = new HashSet<string>(new[] {
        "aria-activedescendant", "aria-controls", "aria-describedby", "aria-details", "aria-errormessage",
        "aria-flowto", "aria-labelledby", "aria-owns"
    }, StringComparer.Ordinal);

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
        foreach (XElement element in root.DescendantsAndSelf()) {
            token.ThrowIfCancellationRequested();
            bool xhtml = element.Name.NamespaceName == "http://www.w3.org/1999/xhtml";
            bool svg = element.Name.NamespaceName == "http://www.w3.org/2000/svg";
            if (!xhtml && !svg) continue;
            foreach (XAttribute attribute in element.Attributes().Where(attribute => attribute.Name.NamespaceName.Length == 0 &&
                (ReferenceAttributes.Contains(attribute.Name.LocalName) ||
                 (xhtml && (element.Name.LocalName == "td" || element.Name.LocalName == "th") && attribute.Name.LocalName == "headers")))) {
                foreach (string id in attribute.Value.Split(new[] { ' ', '\t', '\r', '\n', '\f' }, StringSplitOptions.RemoveEmptyEntries)) {
                    token.ThrowIfCancellationRequested();
                    if (!ids.Contains(id)) throw new InvalidDataException("Document-local " + attribute.Name.LocalName +
                        " target missing in " + path + ": " + id + ". Keep the related content in one chapter or repair the reference.");
                }
            }
        }
    }
}
