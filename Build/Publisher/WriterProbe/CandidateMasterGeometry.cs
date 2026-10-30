using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Constructs the master guide array consumed by native page rendering.</summary>
internal static class CandidateMasterGeometry {
    private const ushort AxisGuideCount = 2;

    internal static byte[] Create(uint width, uint height) => Chunk(
        Scalar(1, 0x20, 4), // CGpv: exactly four allocated guide rows.
        Variable(2, 0xA0, Guide(0), Guide(width), Guide(0), Guide(height)),
        Variable(3, 0x88, Scalar(1, 0x20, width), Scalar(2, 0x20, height)), // DptlPage
        Scalar(4, 0x20, 0), Scalar(5, 0x20, 0), // Zero horizontal/vertical gutters.
        ShortScalar(6, 0x18, AxisGuideCount), // CColGuides: x endpoints precede y endpoints.
        ShortScalar(7, 0x18, AxisGuideCount)); // CRowGuides

    // OplGpv defaults all gutter flags to false. The loader's endpoint
    // normalization adds the declared gutters, which are explicitly zero here.
    private static byte[] Guide(uint coordinate) => Variable(0, 0x88, Scalar(1, 0x20, coordinate));
}
