namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void ThrowIfUnsupportedStyleTableCellProperties(string styleId, LegacyDocElement properties) {
            foreach (LegacyDocElement child in properties.Elements()) {
                if (child.Is("shd")) {
                    ReadSupportedTableCellShading(child, "table style cell shading");
                } else {
                    throw new NotSupportedException($"Native DOC saving supports table style '{styleId}' cell defaults only with shading. Unsupported cell property: {child.LocalName}.");
                }
            }
        }
    }
}
