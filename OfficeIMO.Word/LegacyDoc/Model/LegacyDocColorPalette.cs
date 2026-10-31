namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static class LegacyDocColorPalette {
        private static readonly string?[] IcoToHex = {
            null,
            "000000",
            "0000FF",
            "00FFFF",
            "00FF00",
            "FF00FF",
            "FF0000",
            "FFFF00",
            "FFFFFF",
            "000080",
            "008080",
            "008000",
            "800080",
            "800000",
            "808000",
            "808080",
            "C0C0C0"
        };

        internal static string? GetHexForIco(byte ico) {
            return ico < IcoToHex.Length ? IcoToHex[ico] : null;
        }

        /// <summary>Chooses the closest native RGB palette entry; ties retain the earlier ICO.</summary>
        internal static bool TryGetNearestIcoForHex(string? colorHex, out byte ico) {
            ico = 0;
            string normalized = colorHex?.Replace("#", string.Empty) ?? string.Empty;
            if (normalized.Length != 6 || !uint.TryParse(normalized,
                System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture,
                out uint rgb)) return false;
            int red = (int)(rgb >> 16), green = (int)((rgb >> 8) & 255), blue = (int)(rgb & 255);
            int nearestDistance = int.MaxValue;
            for (byte index = 1; index < IcoToHex.Length; index++) {
                uint candidate = uint.Parse(IcoToHex[index]!, System.Globalization.NumberStyles.HexNumber,
                    System.Globalization.CultureInfo.InvariantCulture);
                int dr = red - (int)(candidate >> 16);
                int dg = green - (int)((candidate >> 8) & 255);
                int db = blue - (int)(candidate & 255);
                int distance = dr * dr + dg * dg + db * db;
                if (distance < nearestDistance) { nearestDistance = distance; ico = index; }
            }
            return ico != 0;
        }

        internal static bool TryGetIcoForHex(string? fillColorHex, out byte ico) {
            ico = 0;
            if (string.IsNullOrWhiteSpace(fillColorHex)) {
                return true;
            }

            string normalized = fillColorHex!.Replace("#", string.Empty).ToUpperInvariant();
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
