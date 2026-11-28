using System.Xml;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Compares projected markup regardless of namespace prefixes, empty-element syntax or attribute order.</summary>
    internal static class XlsbProjectionXmlComparer {
        private const string Compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        private const string SchemaInstance = "http://www.w3.org/2001/XMLSchema-instance";
        private const string Xmlns = "http://www.w3.org/2000/xmlns/";

        internal static bool SameContent(OfficeXmlNode actual, OfficeXmlNode expected) {
            byte[] leftBytes = actual.OuterXml();
            byte[] rightBytes = expected.OuterXml();
            if (leftBytes.AsSpan().SequenceEqual(rightBytes)) return true;
            if (leftBytes.Length == 0 || rightBytes.Length == 0) return false;

            // Read serialized XML because SDK text elements can hide unknown child markup from their DOM.
            using var leftStream = new MemoryStream(leftBytes, writable: false);
            using var rightStream = new MemoryStream(rightBytes, writable: false);
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit };
            using XmlReader left = XmlReader.Create(leftStream, settings);
            using XmlReader right = XmlReader.Create(rightStream, settings);
            try {
                while (left.Read()) {
                    if (!right.Read() || left.NodeType != right.NodeType || left.LocalName != right.LocalName
                        || left.NamespaceURI != right.NamespaceURI || left.Value != right.Value) return false;
                    if (left.NodeType != XmlNodeType.Element) continue;
                    if (!SameAttributes(left, right)) return false;
                    if (left.IsEmptyElement != right.IsEmptyElement) {
                        XmlReader expanded = left.IsEmptyElement ? right : left;
                        if (!expanded.Read() || expanded.NodeType != XmlNodeType.EndElement) return false;
                    }
                }
                return !right.Read();
            } catch (XmlException) {
                // A shadow text element can serialize invalid markup. The native writer must refuse it.
                return false;
            }
        }

        private static bool SameAttributes(XmlReader left, XmlReader right) {
            int count = 0;
            while (left.MoveToNextAttribute()) {
                if (left.NamespaceURI == Xmlns) continue;
                if (left.Value != right.GetAttribute(left.LocalName, left.NamespaceURI)) return false;
                if (left.NamespaceURI == Compatibility || (left.NamespaceURI == SchemaInstance && left.LocalName == "type")
                    || (left.LocalName == "Requires" && left.NamespaceURI.Length == 0)) {
                    // These attribute values can refer to namespaces by their lexical prefix.
                    foreach (string token in left.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
                        int colon = token.IndexOf(':');
                        string prefix = colon < 0 ? token : token.Substring(0, colon);
                        if (left.LookupNamespace(prefix) != right.LookupNamespace(prefix)) return false;
                    }
                }
                count++;
            }
            left.MoveToElement();
            while (right.MoveToNextAttribute()) {
                if (right.NamespaceURI != Xmlns) count--;
            }
            right.MoveToElement();
            return count == 0;
        }
    }
}
