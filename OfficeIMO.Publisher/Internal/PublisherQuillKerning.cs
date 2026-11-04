namespace OfficeIMO.Publisher.Internal;

/// <summary>The qualified native scalar and omitted character-format default, independent of generic drawing defaults.</summary>
internal static class PublisherQuillKerning {
    internal const double DefaultThreshold = 14D;

    /// <summary>Distinguishes an omitted value from an opaque or ambiguous encoded setting.</summary>
    internal static bool TryRead(IReadOnlyList<PublisherBlock> fields, out double? threshold) {
        threshold = null;
        bool found = false;
        foreach (PublisherBlock field in fields) {
            if (field.Id != 0x18) continue;
            if (found || field.Type != 0x22 || field.Value > 12700 * 999.5D) return false;
            found = true;
            threshold = field.Value / 12700D;
        }
        return true;
    }
}
