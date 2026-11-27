using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void ThrowIfUnsupportedTableGapShading(LegacyDocElement properties,
            LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            string? tableStyleId = properties.Child("tblStyle").Val;
            int spacing = ReadSupportedTableDefaultCellSpacingFromProperties(properties)
                ?? ReadSupportedTableStyleDefaultCellSpacing(tableStyleId, tableStyleDefinitions) ?? 0;
            if (spacing <= 0) return;

            LegacyDocElement directShading = properties.Child("shd");
            LegacyDocTableCellShading shading = directShading.Exists
                ? ReadSupportedTableCellShading(directShading, "table gap shading")
                : ReadSupportedTableStyleGapShading(tableStyleId, tableStyleDefinitions);
            if (shading.HasAny) {
                throw new NotSupportedException("Native DOC saving does not support visible table gap shading with positive cell spacing. Remove the gap shading or use zero cell spacing before saving as DOC.");
            }
        }

        private static LegacyDocTableCellShading ReadSupportedTableStyleGapShading(string? tableStyleId,
            LegacyDocTableStyleDefinitions tableStyleDefinitions) {
            LegacyDocElement style = ResolveSupportedTableStyle(tableStyleId, tableStyleDefinitions);
            if (style.IsNone) return default;
            return ReadSupportedTableStyleOwnGapShading(style)
                ?? ReadSupportedTableStyleBaseValue(style, tableStyleDefinitions,
                    new HashSet<string>(StringComparer.Ordinal), ReadSupportedTableStyleOwnGapShading)
                ?? default;
        }

        private static LegacyDocTableCellShading? ReadSupportedTableStyleOwnGapShading(LegacyDocElement style) {
            LegacyDocElement shading = style.Child("tblPr").Child("shd");
            return shading.IsNone ? null : ReadSupportedTableCellShading(shading, "table style gap shading");
        }

    }
}
