namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFaceCollection {
    // Sectioned renderers provide the same immutable byte snapshot for each use of a package font.
    // Reuse the prepared instance without hashing those bytes again in every section.
    internal sealed class OperationFontCache {
        private readonly Dictionary<OperationFontKey, FaceTemplate> _templates = new();
        private long _decodedBytes;

        internal bool TryGet(OperationFontKey key, int maximumDecodedBytes, out OfficeFontFace? face, out int decodedBytes) {
            if (_templates.TryGetValue(key, out FaceTemplate? template) && template.DecodedBytes <= maximumDecodedBytes) {
                face = template.Face.Clone();
                decodedBytes = template.DecodedBytes;
                return true;
            }
            face = null;
            decodedBytes = 0;
            return false;
        }

        internal void Retain(OperationFontKey key, OfficeFontFace face, int decodedBytes) {
            if (decodedBytes > MaximumContentTemplateBytes - _decodedBytes) return;
            if (_templates.ContainsKey(key)) return;
            _templates.Add(key, new FaceTemplate(face.Clone(), decodedBytes));
            _decodedBytes += decodedBytes;
        }
    }

    internal readonly record struct OperationFontKey(byte[] Data, string Family, OfficeFontFaceDescriptor Descriptor, OfficeFontUnicodeRangeSet Ranges);

    internal bool TryAddOperationFont(OperationFontCache cache, string family, byte[] data, OfficeFontFaceDescriptor descriptor,
        OfficeFontUnicodeRangeSet ranges, int maximumDecodedBytes, out int decodedBytes, out string? error) {
        if (string.IsNullOrWhiteSpace(family) || data.Length == 0)
            return TryAddContentBounded(family, data, descriptor, ranges, maximumDecodedBytes, out decodedBytes, out error);
        var key = new OperationFontKey(data, family.Trim(), descriptor, ranges);
        if (cache.TryGet(key, maximumDecodedBytes, out OfficeFontFace? cached, out decodedBytes)) {
            ReplaceOrAddFace(cached!);
            error = null;
            return true;
        }

        if (!TryAddContentBounded(family, data, descriptor, ranges, maximumDecodedBytes, out decodedBytes, out error))
            return false;
        string resourceFamily = descriptor == OfficeFontFaceDescriptor.Regular && ranges.IsAll
            ? key.Family
            : CreateResourceFamilyName(key.Family, descriptor, ranges);
        OfficeFontFace? added = _faces.FindLast(face => face.ResourceFamilyName == resourceFamily && face.Descriptor == descriptor);
        if (added != null) cache.Retain(key, added, decodedBytes);
        return true;
    }
}
