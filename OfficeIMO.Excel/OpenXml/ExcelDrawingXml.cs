using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel.OpenXml {
    /// <summary>
    /// Element helpers worksheet drawing and image code uses over <see cref="OfficeXmlNode"/>: child orders the generated
    /// views don't cover, walking descendants by local name in any namespace, and reading relationship attributes.
    /// </summary>
    internal static class ExcelDrawingXml {
        private const string RelationshipsUri = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        /// <summary>The children of <c>xdr:from</c> and <c>xdr:to</c> (CT_Marker), in schema order.</summary>
        internal static readonly OfficeXmlChildOrder MarkerOrder = new(
            new[] { ExcelChartXml.Xdr, ExcelChartXml.Xdr, ExcelChartXml.Xdr, ExcelChartXml.Xdr },
            new[] { "col", "colOff", "row", "rowOff" },
            new[] { 0, 1, 2, 3 });

        /// <summary>The children of <c>xdr:cNvPr</c> (CT_NonVisualDrawingProps), in schema order.</summary>
        internal static readonly OfficeXmlChildOrder NonVisualDrawingPropertiesOrder = new(
            new[] { ExcelChartXml.A, ExcelChartXml.A, ExcelChartXml.A },
            new[] { "hlinkClick", "hlinkHover", "extLst" },
            new[] { 0, 1, 2 });

        /// <summary>The descendants with this local name in any namespace, in document order.</summary>
        internal static List<OfficeXmlNode> DescendantsNamed(OfficeXmlNode element, string localName) {
            var found = new List<OfficeXmlNode>();
            if (!element.IsNone) CollectNamed(element, localName, found);
            return found;
        }

        private static void CollectNamed(OfficeXmlNode element, string localName, List<OfficeXmlNode> found) {
            for (OfficeXmlNode child = element.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (string.Equals(child.LocalName, localName, StringComparison.Ordinal)) found.Add(child);
                CollectNamed(child, localName, found);
            }
        }

        /// <summary>The first descendant with this local name in any namespace, in document order, or none.</summary>
        internal static OfficeXmlNode FirstDescendantNamed(OfficeXmlNode element, string localName) {
            if (element.IsNone) return default;
            for (OfficeXmlNode child = element.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (string.Equals(child.LocalName, localName, StringComparison.Ordinal)) return child;
                OfficeXmlNode found = FirstDescendantNamed(child, localName);
                if (!found.IsNone) return found;
            }
            return default;
        }

        /// <summary>The first descendant with this name in a namespace outside the fixed ids, or none.</summary>
        internal static OfficeXmlNode FirstDescendant(OfficeXmlNode element, string namespaceUri, string localName) {
            if (element.IsNone) return default;
            for (OfficeXmlNode child = element.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (child.Is(namespaceUri, localName)) return child;
                OfficeXmlNode found = FirstDescendant(child, namespaceUri, localName);
                if (!found.IsNone) return found;
            }
            return default;
        }

        /// <summary>The descendants with this name in a namespace outside the fixed ids, in document order.</summary>
        internal static List<OfficeXmlNode> Descendants(OfficeXmlNode element, string namespaceUri, string localName) {
            var found = new List<OfficeXmlNode>();
            if (!element.IsNone) CollectNamespaced(element, namespaceUri, localName, found);
            return found;
        }

        private static void CollectNamespaced(OfficeXmlNode element, string namespaceUri, string localName, List<OfficeXmlNode> found) {
            for (OfficeXmlNode child = element.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (child.Is(namespaceUri, localName)) found.Add(child);
                CollectNamespaced(child, namespaceUri, localName, found);
            }
        }

        /// <summary>
        /// Adds the non-empty values of every relationship-namespace attribute on the element and its descendants, such
        /// as <c>r:embed</c>, <c>r:id</c> and the references extensions keep.
        /// </summary>
        internal static void CollectRelationshipAttributeValues(OfficeXmlNode element, ISet<string> values) {
            ForEachRelationshipAttribute(element, (_, value) => {
                if (!string.IsNullOrWhiteSpace(value)) values.Add(value);
                return false;
            });
        }

        /// <summary>Whether the element or a descendant has a relationship-namespace attribute with this name and value.</summary>
        internal static bool HasRelationshipAttribute(OfficeXmlNode element, string localName, string value) =>
            ForEachRelationshipAttribute(element, (name, attributeValue) =>
                string.Equals(name, localName, StringComparison.Ordinal) && string.Equals(attributeValue, value, StringComparison.Ordinal));

        // Visits the element and its descendants in document order. Stops when the visitor returns true.
        private static bool ForEachRelationshipAttribute(OfficeXmlNode element, Func<string, string, bool> visit) {
            if (element.IsNone) return false;
            if (element.HasAttributes) {
                foreach (OfficeXmlAttribute attribute in element.Attributes()) {
                    if (string.Equals(attribute.NamespaceUri, RelationshipsUri, StringComparison.Ordinal)
                        && visit(attribute.LocalName, attribute.Value)) return true;
                }
            }
            for (OfficeXmlNode child = element.FirstChild; !child.IsNone; child = child.NextSibling) {
                if (ForEachRelationshipAttribute(child, visit)) return true;
            }
            return false;
        }
    }
}
