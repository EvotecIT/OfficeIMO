using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        /// <summary>Removes content and the image and chart parts only it referenced.</summary>
        internal static void RemoveWithOwnedParts(OpenXmlElement element) {
            string[] ids = RelationshipIds(element).Distinct(StringComparer.Ordinal).ToArray();
            OpenXmlPartRootElement? root = element.Ancestors<OpenXmlPartRootElement>().FirstOrDefault();
            element.Remove();
            if (ids.Length == 0 || root?.OpenXmlPart is not { } part) return;

            var remaining = new HashSet<string>(RelationshipIds(root), StringComparer.Ordinal);
            foreach (string id in ids) {
                if (remaining.Contains(id)) continue;
                IdPartPair pair = part.Parts.FirstOrDefault(candidate => candidate.RelationshipId == id);
                if (pair.OpenXmlPart is ImagePart or ChartPart) part.DeletePart(id);
            }
        }

        private static IEnumerable<string> RelationshipIds(OpenXmlElement element) =>
            new[] { element }.Concat(element.Descendants())
                .SelectMany(node => node.GetAttributes())
                .Where(attribute => attribute.NamespaceUri == RelationshipNamespace && attribute.Value != null)
                .Select(attribute => attribute.Value!);
    }
}
