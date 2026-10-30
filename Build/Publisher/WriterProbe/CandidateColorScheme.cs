using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Constructs the publication's explicit, eight-slot RGB color scheme.</summary>
internal static class CandidateColorScheme {
    internal static byte[] CreateDefault() {
        byte[][] colors = {
            Marker(0, 0x78), // Automatic black, using the native compressed row.
            Color(91, 155, 213), // Blue.
            Color(237, 125, 49), // Orange.
            Color(158, 158, 158), // Gray.
            Color(255, 192, 0), // Gold.
            Color(8, 82, 150), // Dark blue.
            Color(153, 51, 102), // Purple.
            Color(255, 255, 255) // White.
        };
        // OplSccm: Cecp and Rgecp. Optional scheme identity/name and extended
        // color properties retain their qualified native defaults.
        return Chunk(Scalar(1, 0x20, (uint)colors.Length), Variable(2, 0xA0, colors));
    }

    private static byte[] Color(byte red, byte green, byte blue) =>
        Variable(0, 0x88, Scalar(1, 0x20, (uint)(red | (green << 8) | (blue << 16))));
}
