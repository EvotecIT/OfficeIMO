namespace OfficeIMO.Drawing;

/// <summary>Recognizes Windows DIB header layouts supported by the raster and icon readers.</summary>
internal static class OfficeDibHeaderLayout {
    internal static bool IsSupportedWindowsInfoHeaderSize(int headerSize) =>
        headerSize == 40 || headerSize == 52 || headerSize == 56 ||
        headerSize == 108 || headerSize == 124;
}
