namespace OfficeIMO.Drawing;

/// <summary>Shared cmap platform and encoding classification.</summary>
internal static class OfficeOpenTypeCmap {
    internal static bool IsUnicodeEncoding(int platform, int encoding) =>
        platform == 0 ||
        platform == 3 && (encoding == 1 || encoding == 10);
}
