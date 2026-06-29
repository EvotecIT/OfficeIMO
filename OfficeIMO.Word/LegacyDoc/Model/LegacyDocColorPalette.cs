namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocColorPalette {
        private static readonly string?[] IcoToHex = {
            null,
            "000000",
            "0000ff",
            "00ffff",
            "00ff00",
            "ff00ff",
            "ff0000",
            "ffff00",
            "ffffff",
            "000080",
            "008080",
            "008000",
            "800080",
            "800000",
            "808000",
            "808080",
            "c0c0c0"
        };

        internal static string? GetHexForIco(byte ico) {
            return ico < IcoToHex.Length ? IcoToHex[ico] : null;
        }

        internal static bool TryGetIcoForHex(string? fillColorHex, out byte ico) {
            ico = 0;
            if (string.IsNullOrWhiteSpace(fillColorHex)) {
                return true;
            }

            string normalized = fillColorHex!.Replace("#", string.Empty).ToLowerInvariant();
            for (byte index = 1; index < IcoToHex.Length; index++) {
                if (string.Equals(IcoToHex[index], normalized, StringComparison.Ordinal)) {
                    ico = index;
                    return true;
                }
            }

            return false;
        }
    }
}
