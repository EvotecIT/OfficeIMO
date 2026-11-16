namespace OfficeIMO.Word.LegacyDoc.Model {
    internal readonly struct LegacyDocParagraphShading : IEquatable<LegacyDocParagraphShading> {
        internal LegacyDocParagraphShading(string? fillColorHex) {
            FillColorHex = string.IsNullOrWhiteSpace(fillColorHex)
                ? null
                : fillColorHex!.Replace("#", string.Empty).ToUpperInvariant();
        }

        internal string? FillColorHex { get; }

        internal bool HasAny => !string.IsNullOrEmpty(FillColorHex);

        internal bool IsAutomatic => string.Equals(FillColorHex, "AUTO", StringComparison.Ordinal);

        /// <summary>The fill in OpenXML notation: six-digit RGB or lowercase <c>auto</c>.</summary>
        internal string? OpenXmlFill => IsAutomatic ? "auto" : FillColorHex;

        internal bool TryGetRgb(out byte red, out byte green, out byte blue) {
            red = green = blue = 0;
            string? fill = FillColorHex;
            return fill != null && fill.Length == 6
                && byte.TryParse(fill.Substring(0, 2), System.Globalization.NumberStyles.AllowHexSpecifier, System.Globalization.CultureInfo.InvariantCulture, out red)
                && byte.TryParse(fill.Substring(2, 2), System.Globalization.NumberStyles.AllowHexSpecifier, System.Globalization.CultureInfo.InvariantCulture, out green)
                && byte.TryParse(fill.Substring(4, 2), System.Globalization.NumberStyles.AllowHexSpecifier, System.Globalization.CultureInfo.InvariantCulture, out blue);
        }

        public bool Equals(LegacyDocParagraphShading other) {
            return string.Equals(FillColorHex, other.FillColorHex, StringComparison.Ordinal);
        }

        public override bool Equals(object? obj) {
            return obj is LegacyDocParagraphShading other && Equals(other);
        }

        public override int GetHashCode() {
            return FillColorHex == null ? 0 : FillColorHex.GetHashCode();
        }
    }
}
