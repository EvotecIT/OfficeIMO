namespace OfficeIMO.Word.LegacyDoc.Write {
    /// <summary>
    /// The style definitions of a styles part: the top-level <c>w:style</c> elements in document order, lookup by
    /// identifier (the first definition wins) and the document defaults.
    /// </summary>
    internal sealed class LegacyDocStyleCatalog {
        private readonly Dictionary<string, LegacyDocElement> _firstById = new(StringComparer.Ordinal);

        internal LegacyDocStyleCatalog(LegacyDocElement styles) {
            Root = styles;
            var definitions = new List<LegacyDocElement>();
            foreach (LegacyDocElement style in styles.Elements("style")) {
                definitions.Add(style);
                string? id = style.Attr("styleId");
                if (id != null && !_firstById.ContainsKey(id)) _firstById.Add(id, style);
            }
            Definitions = definitions;
        }

        /// <summary>The <c>w:styles</c> element, or none when the document has no styles part.</summary>
        internal LegacyDocElement Root { get; }

        internal bool Exists => Root.Exists;

        internal IReadOnlyList<LegacyDocElement> Definitions { get; }

        /// <summary>The <c>w:docDefaults</c> element, or none.</summary>
        internal LegacyDocElement DocumentDefaults => Root.Child("docDefaults");

        /// <summary>The first style with this identifier, whatever its type, or none.</summary>
        internal LegacyDocElement FindById(string? styleId) =>
            styleId != null && _firstById.TryGetValue(styleId, out LegacyDocElement style) ? style : LegacyDocElement.None;

        internal static string? StyleId(LegacyDocElement style) => style.Attr("styleId");

        internal static string? Type(LegacyDocElement style) => style.Attr("type");

        internal static bool IsParagraphStyle(LegacyDocElement style) => string.Equals(style.Attr("type"), "paragraph", StringComparison.Ordinal);

        internal static bool IsDefault(LegacyDocElement style) => style.Attr("default") is { } value && LegacyDocValues.ParseOnOff(value);
    }
}
