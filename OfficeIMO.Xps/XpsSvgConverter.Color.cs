using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private readonly Dictionary<string, OfficeIccColorProfile?> _profiles = new(StringComparer.OrdinalIgnoreCase);
    private long _profileAllowance;
    private static double Clamp(double value) => Math.Max(0, Math.Min(1, value));

    private OfficeIccColorProfile? ColorProfile(string part, string uri, bool reportUnsupported = true) {
        _token.ThrowIfCancellationRequested();
        string name = XpsPackage.Resolve(part, uri);
        if (_profiles.TryGetValue(name, out var existing)) {
            if (existing == null && reportUnsupported) Loss("Unsupported ICC profile");
            return existing;
        }
        if (_page.Document.ContentType(name) != "application/vnd.ms-color.iccprofile") throw new InvalidDataException("Invalid ICC resource content type.");
        return ParseColorProfile(name, _page.Document.Part(name), reportUnsupported);
    }

    private OfficeIccColorProfile? ParseColorProfile(string name, byte[] bytes, bool reportUnsupported = true) {
        if (_profiles.TryGetValue(name, out var existing)) {
            if (existing == null && reportUnsupported) Loss("Unsupported ICC profile");
            return existing;
        }
        // LUT parsing expands encoded samples. Reserve the same conservative parser
        // allowance used by Core's packed ICC raster converter before parsing each profile.
        long allowance = bytes.LongLength * 32L + 4096;
        if (bytes.Length > 4 * 1024 * 1024 || allowance > 64 * 1024 * 1024 - _profileAllowance)
            throw new InvalidDataException("XPS ICC profile budget exceeded.");
        _profileAllowance += allowance;
        if (!OfficeIccColorProfile.TryCreate(bytes, out var profile) && reportUnsupported) Loss("Unsupported ICC profile");
        _token.ThrowIfCancellationRequested();
        if (profile != null && profile.RetainedByteCount > allowance) {
            long extra = profile.RetainedByteCount - allowance;
            if (extra > 64 * 1024 * 1024 - _profileAllowance) throw new InvalidDataException("XPS ICC profile budget exceeded.");
            _profileAllowance += extra;
        }
        _profiles.Add(name, profile);
        return profile;
    }
}
