namespace OfficeIMO.Drawing;

public sealed partial class OfficeFontFaceCollection {
    // Released HTML loaders use this exact descriptor overload. Descriptor weight
    // remains opt-in for newer callers; older callers keep the font's default instance.
    internal bool TryAddBounded(
        string? familyName,
        byte[]? data,
        OfficeFontFaceDescriptor descriptor,
        OfficeFontUnicodeRangeSet? unicodeRanges,
        int maximumDecodedBytes,
        out int decodedBytes,
        out string? error) =>
        TryAddBounded(familyName, data, descriptor, unicodeRanges, maximumDecodedBytes,
            out decodedBytes, out error, applyDescriptorWeight: false);
}
