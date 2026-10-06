using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void ThrowIfUnsupportedStyleTableCellProperties(string styleId, StyleTableCellProperties properties) {
            foreach (OpenXmlElement child in properties.ChildElements) {
                if (child is Shading shading) {
                    ReadSupportedTableCellShading(shading, "table style cell shading");
                } else {
                    throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' cell defaults only with shading. Unsupported cell property: {child.LocalName}.");
                }
            }
        }
    }
}
