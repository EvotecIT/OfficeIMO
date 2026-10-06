using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.LegacyDoc.Model;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private static void ThrowIfUnsupportedTableGapShading(TableProperties properties,
            IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            TableStyle? tableStyle = properties.GetFirstChild<TableStyle>();
            int spacing = ReadSupportedTableDefaultCellSpacing(properties)
                ?? ReadSupportedTableStyleDefaultCellSpacing(tableStyle, tableStyleDefinitions) ?? 0;
            if (spacing <= 0) return;

            Shading? directShading = properties.GetFirstChild<Shading>();
            LegacyDocTableCellShading shading = directShading != null
                ? ReadSupportedTableCellShading(directShading, "table gap shading")
                : ReadSupportedTableStyleGapShading(tableStyle, tableStyleDefinitions);
            if (shading.HasAny) {
                throw new NotSupportedException("Native DOC saving does not support visible table gap shading with positive cell spacing. Remove the gap shading or use zero cell spacing before saving as DOC.");
            }
        }

        private static LegacyDocTableCellShading ReadSupportedTableStyleGapShading(TableStyle? tableStyle,
            IReadOnlyDictionary<string, Style> tableStyleDefinitions) {
            Style? style = ResolveSupportedTableStyle(tableStyle, tableStyleDefinitions);
            if (style == null) return default;
            return ReadSupportedTableStyleOwnGapShading(style)
                ?? ReadSupportedTableStyleBaseValue(style, tableStyleDefinitions,
                    new HashSet<string>(StringComparer.OrdinalIgnoreCase), ReadSupportedTableStyleOwnGapShading)
                ?? default;
        }

        private static LegacyDocTableCellShading? ReadSupportedTableStyleOwnGapShading(Style style) {
            Shading? shading = style.GetFirstChild<StyleTableProperties>()?.GetFirstChild<Shading>();
            return shading == null ? null : ReadSupportedTableCellShading(shading, "table style gap shading");
        }

    }
}
