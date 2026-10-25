namespace OfficeIMO.Html;

internal static partial class HtmlCssBoxStrokeParser {
    private static readonly Lazy<string[]> StrokeProperties = new(() => new[] {
        "border", "border-width", "border-style", "border-color"
    }.Concat(SideProperties).Concat(new[] {
        "outline", "outline-width", "outline-style", "outline-color", "outline-offset"
    }).ToArray());

    /// <summary>Captures immutable declarations and cascade state used by border and outline parsing.</summary>
    /// <remarks>Numeric context and current color belong to the resolver's separate used-value key.</remarks>
    internal sealed class StrokeDeclarations : IEquatable<StrokeDeclarations> {
        private readonly Entry[] _entries;
        private readonly int _hashCode;

        internal StrokeDeclarations(HtmlComputedStyle computed) {
            var entries = new List<Entry>();
            int hash = 17;
            string[] properties = StrokeProperties.Value;
            for (int index = 0; index < properties.Length; index++) {
                string property = properties[index];
                if (!computed.HasCascadeState(property)) continue;
                byte state = 0;
                if (computed.IsInheritedValue(property)) state |= 1;
                if (computed.IsResetValue(property)) state |= 2;
                if (computed.IsOriginRevertedValue(property)) state |= 4;
                if (computed.IsSpecifiedValue(property)) state |= 8;
                HtmlCssCascadePriority? priority = computed.TryGetCascadePriority(property, out var found)
                    ? found : null;
                var entry = new Entry(index, computed.GetValue(property), state, priority);
                entries.Add(entry);
                hash = unchecked(hash * 31 + entry.GetHashCode());
            }
            _entries = entries.ToArray();
            _hashCode = hash;
        }

        public bool Equals(StrokeDeclarations? other) {
            if (ReferenceEquals(this, other)) return true;
            if (other == null || _hashCode != other._hashCode || _entries.Length != other._entries.Length) return false;
            for (int index = 0; index < _entries.Length; index++) {
                if (!_entries[index].Equals(other._entries[index])) return false;
            }
            return true;
        }

        public override bool Equals(object? obj) => obj is StrokeDeclarations other && Equals(other);

        public override int GetHashCode() => _hashCode;

        private readonly record struct Entry(int Property, string Value, byte State, HtmlCssCascadePriority? Priority);
    }
}
