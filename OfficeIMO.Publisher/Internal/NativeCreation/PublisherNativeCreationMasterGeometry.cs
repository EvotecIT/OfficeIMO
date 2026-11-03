
namespace OfficeIMO.Publisher.Internal;

/// <summary>Constructs the master guide array consumed by native page rendering.</summary>
internal static class PublisherNativeCreationMasterGeometry {
    private const ushort AxisGuideCount = 2;

    internal static byte[] Create(uint width, uint height) => PublisherOplWriter.Chunk(
        PublisherOplWriter.UInt32(1, 4, 0x20), // CGpv: exactly four allocated guide rows.
        PublisherOplWriter.Nested(2, 0xA0, Guide(0), Guide(width), Guide(0), Guide(height)),
        PublisherOplWriter.Nested(3, 0x88, PublisherOplWriter.UInt32(1, width, 0x20), PublisherOplWriter.UInt32(2, height, 0x20)), // DptlPage
        PublisherOplWriter.UInt32(4, 0, 0x20), PublisherOplWriter.UInt32(5, 0, 0x20), // Zero horizontal/vertical gutters.
        PublisherOplWriter.UInt16(6, AxisGuideCount, 0x18), // CColGuides: x endpoints precede y endpoints.
        PublisherOplWriter.UInt16(7, AxisGuideCount, 0x18)); // CRowGuides

    // OplGpv defaults all gutter flags to false. The loader's endpoint
    // normalization adds the declared gutters, which are explicitly zero here.
    private static byte[] Guide(uint coordinate) => PublisherOplWriter.Nested(0, 0x88, PublisherOplWriter.UInt32(1, coordinate, 0x20));
}
