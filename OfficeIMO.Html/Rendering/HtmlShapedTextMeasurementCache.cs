using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Operation-scoped bounded cache for successful configured-shaper measurements.</summary>
internal sealed class HtmlShapedTextMeasurementCache {
    internal const int MaximumEntries = 4096;
    internal const int MaximumRetainedKeyCharacters = 256 * 1024;
    internal const int MaximumCacheableTextCharacters = 1024;

    private readonly Dictionary<Key, double> _values = new Dictionary<Key, double>();
    private int _retainedKeyCharacters;

    internal int Count => _values.Count;
    internal int RetainedKeyCharacters => _retainedKeyCharacters;

    internal bool TryGet(string text, OfficeFontInfo font, OfficeTextFeatureSettings features, string? language, out double measured) =>
        _values.TryGetValue(new Key(text, font, features, language), out measured);

    internal void Store(string text, OfficeFontInfo font, OfficeTextFeatureSettings features, string? language, double measured) {
        if (text.Length == 0
            || text.Length > MaximumCacheableTextCharacters
            || _values.Count >= MaximumEntries
            || text.Length > MaximumRetainedKeyCharacters - _retainedKeyCharacters
            || (language?.Length ?? 0) > MaximumRetainedKeyCharacters - _retainedKeyCharacters - text.Length) {
            return;
        }

        var key = new Key(text, font, features, language);
        if (_values.ContainsKey(key)) return;
        _values.Add(key, measured);
        _retainedKeyCharacters += text.Length + (language?.Length ?? 0);
    }

    private readonly struct Key : IEquatable<Key> {
        internal Key(string text, OfficeFontInfo font, OfficeTextFeatureSettings features, string? language) {
            Text = text;
            Font = font;
            Features = features;
            Language = language;
        }

        private string Text { get; }
        private OfficeFontInfo Font { get; }
        private OfficeTextFeatureSettings Features { get; }
        private string? Language { get; }

        public bool Equals(Key other) =>
            string.Equals(Text, other.Text, StringComparison.Ordinal) && Font.Equals(other.Font) && Features.Equals(other.Features)
            && string.Equals(Language, other.Language, StringComparison.Ordinal);

        public override bool Equals(object? obj) => obj is Key other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                int hash = ((StringComparer.Ordinal.GetHashCode(Text) * 397) ^ Font.GetHashCode()) * 397 ^ Features.GetHashCode();
                return hash * 397 ^ (Language == null ? 0 : StringComparer.Ordinal.GetHashCode(Language));
            }
        }
    }
}
