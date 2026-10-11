using System.Xml;

namespace OfficeIMO.Excel {
    /// <summary>Matches owned SpreadsheetML content in transitional and strict worksheet and shared-string XML.</summary>
    internal static class SpreadsheetXmlContent {
        internal const string SpreadsheetNamespace =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        internal const string StrictSpreadsheetNamespace =
            "http://purl.oclc.org/ooxml/spreadsheetml/main";

        /// <summary>Matches an element name only in a supported SpreadsheetML namespace.</summary>
        internal static bool IsSpreadsheetElement(XmlReader reader, string localName) {
            return reader.NodeType == XmlNodeType.Element
                && reader.LocalName == localName
                && (reader.NamespaceURI == SpreadsheetNamespace || reader.NamespaceURI == StrictSpreadsheetNamespace);
        }

        /// <summary>Matches a direct SpreadsheetML child without accepting descendants owned by an extension.</summary>
        internal static bool IsDirectChildElement(XmlReader reader, int parentDepth, string localName) {
            return reader.Depth == parentDepth + 1 && IsSpreadsheetElement(reader, localName);
        }

        /// <summary>Includes direct visible text and direct rich-run text, excluding extension and phonetic payloads.</summary>
        internal static bool IsRichTextElement(XmlReader reader, int ownerDepth, ref int runDepth) {
            if (reader.NodeType == XmlNodeType.EndElement && reader.Depth == runDepth) {
                runDepth = -1;
                return false;
            }

            if (IsDirectChildElement(reader, ownerDepth, reader.LocalName)) {
                if (reader.LocalName == "r") {
                    runDepth = reader.IsEmptyElement ? -1 : reader.Depth;
                    return false;
                }

                return reader.LocalName == "t";
            }

            return runDepth >= 0 && IsDirectChildElement(reader, runDepth, "t");
        }
    }
}
