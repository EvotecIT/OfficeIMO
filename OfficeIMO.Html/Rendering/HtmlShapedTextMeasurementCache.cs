using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Operation-scoped bounded cache for successful configured-shaper measurements.</summary>
internal sealed class HtmlShapedTextMeasurementCache {
    internal const int MaximumEntries = 1024;
    internal const int MaximumRetainedTextCharacters = 256 * 1024;
    internal const int MaximumCacheableTextCharacters = 1024;

    private readonly Dictionary<Key, double> _values = new Dictionary<Key, double>();
    private int _retainedTextCharacters;

    internal int Count => _values.Count;
    internal int RetainedTextCharacters => _retainedTextCharacters;

    internal bool TryGet(string text, OfficeFontInfo font, OfficeTextFeatureSettings features, out double measured) =>
        _values.TryGetValue(new Key(text, font, features), out measured);

    internal bool TryGet(string text, OfficeFontInfo font, out double measured) =>
        TryGet(text, font, OfficeTextFeatureSettings.Default, out measured);

    internal void Store(string text, OfficeFontInfo font, OfficeTextFeatureSettings features, double measured) {
        if (text.Length == 0
            || text.Length > MaximumCacheableTextCharacters
            || _values.Count >= MaximumEntries
            || text.Length > MaximumRetainedTextCharacters - _retainedTextCharacters) {
            return;
        }

        var key = new Key(text, font, features);
        if (_values.ContainsKey(key)) return;
        _values.Add(key, measured);
        _retainedTextCharacters += text.Length;
    }

    internal void Store(string text, OfficeFontInfo font, double measured) =>
        Store(text, font, OfficeTextFeatureSettings.Default, measured);

    private readonly struct Key : IEquatable<Key> {
        internal Key(string text, OfficeFontInfo font, OfficeTextFeatureSettings features) {
            Text = text;
            Font = font;
            Features = features;
        }

        private string Text { get; }
        private OfficeFontInfo Font { get; }
        private OfficeTextFeatureSettings Features { get; }

        public bool Equals(Key other) =>
            string.Equals(Text, other.Text, StringComparison.Ordinal) && Font.Equals(other.Font) && Features.Equals(other.Features);

        public override bool Equals(object? obj) => obj is Key other && Equals(other);

        public override int GetHashCode() {
            unchecked {
                return ((StringComparer.Ordinal.GetHashCode(Text) * 397) ^ Font.GetHashCode()) * 397 ^ Features.GetHashCode();
            }
        }
    }
}
