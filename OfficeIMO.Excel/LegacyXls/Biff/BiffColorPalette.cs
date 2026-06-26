namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffColorPalette {
        private static readonly string[] BuiltInColors = {
            "FF000000", "FFFFFFFF", "FFFF0000", "FF00FF00",
            "FF0000FF", "FFFFFF00", "FFFF00FF", "FF00FFFF"
        };

        private static readonly string[] DefaultPaletteColors = {
            "FF000000", "FFFFFFFF", "FFFF0000", "FF00FF00", "FF0000FF", "FFFFFF00", "FFFF00FF", "FF00FFFF",
            "FF800000", "FF008000", "FF000080", "FF808000", "FF800080", "FF008080", "FFC0C0C0", "FF808080",
            "FF9999FF", "FF993366", "FFFFFFCC", "FFCCFFFF", "FF660066", "FFFF8080", "FF0066CC", "FFCCCCFF",
            "FF000080", "FFFF00FF", "FFFFFF00", "FF00FFFF", "FF800080", "FF800000", "FF008080", "FF0000FF",
            "FF00CCFF", "FFCCFFFF", "FFCCFFCC", "FFFFFF99", "FF99CCFF", "FFFF99CC", "FFCC99FF", "FFFFCC99",
            "FF3366FF", "FF33CCCC", "FF99CC00", "FFFFCC00", "FFFF9900", "FFFF6600", "FF666699", "FF969696",
            "FF003366", "FF339966", "FF003300", "FF333300", "FF993300", "FF993366", "FF333399", "FF333333"
        };

        internal static bool TryResolve(ushort colorIndex, IReadOnlyList<string> customPaletteColors, out string? argb) {
            argb = null;
            if (colorIndex == 0x0051 || colorIndex == 0x7FFF) {
                return false;
            }

            if (colorIndex < BuiltInColors.Length) {
                argb = BuiltInColors[colorIndex];
                return true;
            }

            if (colorIndex >= 0x0008 && colorIndex <= 0x003F) {
                int paletteIndex = colorIndex - 0x0008;
                if (paletteIndex < customPaletteColors.Count) {
                    argb = customPaletteColors[paletteIndex];
                    return true;
                }

                argb = DefaultPaletteColors[paletteIndex];
                return true;
            }

            return false;
        }
    }
}
