using System.Text;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>
/// Synthesizes the observed typed header-property and class-catalog envelopes.
/// Numeric producer traits are explicit hypotheses, not qualified format rules.
/// </summary>
internal static class NativeSchemaPrefix {
    private const uint ObservedFileVersion = 0x000A03E8;
    private const uint ObservedHeaderMarker = 0x0100001A;
    private const uint ObservedPropertyVersion = 0x03E8001A;
    private const uint ObservedClassVersion = 0x03E80000;
    private const uint ObservedProducerBuild = 20527;
    private const int FixedHeaderBytes = 44;

    internal static byte[] Create(IReadOnlyList<uint> objectKinds, uint documentId, int streamLength, int trailerOffset) {
        var classes = new List<ClassDescriptor> {
            new(0x43, Marker6: true, Marker7: true, Trait9: 1, Trait12: 5),
            new(0x44, Marker6: true, Marker7: true, Trait9: 1, Trait12: 5),
            new(0x4B, Marker6: false, Marker7: false, Trait9: 3, Trait12: 5),
            new(0x4C, Marker6: false, Marker7: false, Trait9: 1, Trait12: 5),
            new(0x5B, Marker6: false, Marker7: false, Trait9: null, Trait12: 5),
            new(0x5C, Marker6: false, Marker7: false, Trait9: 1, Trait12: 5),
            new(0x61, Marker6: true, Marker7: false, Trait9: 2, Trait12: 5),
            new(0x65, Marker6: false, Marker7: false, Trait9: 2, Trait12: 5)
        };
        foreach (uint kind in objectKinds) classes.Add(kind == 0x20
            ? new ClassDescriptor(0x20, Marker6: false, Marker7: false, Trait9: null, Trait12: 3)
            : new ClassDescriptor(1, Marker6: true, Marker7: false, Trait9: 1, Trait12: 1));
        byte[] catalog = Chunk(Scalar(1, 0x20, (uint)classes.Count), Scalar(2, 0x20, ObservedPropertyVersion),
            Variable(3, 0x90, classes.OrderBy(item => item.Id).Select(Class).ToArray()));
        int propertiesEnd = FixedHeaderBytes + Properties(documentId, 0, 0, 0).Length;
        byte[] properties = Properties(documentId, streamLength, trailerOffset, propertiesEnd);
        byte[] header = new byte[FixedHeaderBytes];
        header[0] = 0xE8; header[1] = 0xAC; header[2] = 0x2C;
        Put(4, ObservedFileVersion);
        Put(8, (uint)streamLength);
        Put(12, ObservedHeaderMarker);
        Put(0x1A, (uint)trailerOffset);
        Put(0x1E, (uint)propertiesEnd);
        // The two observed 16-bit values are retained without assigning an
        // allocator/packing meaning that the native oracle has not established.
        Put(0x22, 0x00100010);
        return Join(header, properties, catalog);

        void Put(int offset, uint value) => BitConverter.GetBytes(value).CopyTo(header, offset);
    }

    private static byte[] Properties(uint documentId, int streamLength, int trailerOffset, int propertiesEnd) => Chunk(
        ShortScalar((byte)PropertyField.Magic, 0x18, 0xACE8),
        Scalar((byte)PropertyField.DocumentId, 0x68, documentId),
        Scalar((byte)PropertyField.Version3, 0x20, ObservedPropertyVersion),
        Scalar((byte)PropertyField.Version4, 0x20, ObservedPropertyVersion),
        Scalar((byte)PropertyField.Version5, 0x20, ObservedPropertyVersion),
        Scalar((byte)PropertyField.Version7, 0x20, ObservedPropertyVersion),
        Scalar((byte)PropertyField.StreamLength, 0x20, (uint)streamLength),
        Variable((byte)PropertyField.FileName, 0xC0, Encoding.Unicode.GetBytes("WriterProbe.pub\0")),
        Scalar((byte)PropertyField.TrailerOffset, 0xB8, (uint)trailerOffset),
        Scalar((byte)PropertyField.CatalogOffset, 0xB8, (uint)propertiesEnd),
        ShortScalar((byte)PropertyField.ObservedMarker12, 0x18, 1),
        Scalar((byte)PropertyField.Producer13, 0x20, ObservedProducerBuild),
        Scalar((byte)PropertyField.Producer14, 0x20, ObservedProducerBuild),
        Scalar((byte)PropertyField.Producer15, 0x20, ObservedProducerBuild),
        Scalar((byte)PropertyField.Producer16, 0x20, ObservedProducerBuild));

    private static byte[] Class(ClassDescriptor descriptor) {
        var fields = new List<byte[]> {
            ShortScalar((byte)ClassField.Id, 0x18, descriptor.Id),
            Scalar((byte)ClassField.Version, 0x20, ObservedClassVersion)
        };
        if (descriptor.Marker6) fields.Add(Marker((byte)ClassField.ObservedMarker6, 0x08));
        if (descriptor.Marker7) fields.Add(Marker((byte)ClassField.ObservedMarker7, 0x08));
        fields.Add(Marker((byte)ClassField.ObservedMarker8, 0x08));
        if (descriptor.Trait9.HasValue) fields.Add(ShortScalar((byte)ClassField.ObservedTrait9, 0x10, descriptor.Trait9.Value));
        fields.Add(Scalar((byte)ClassField.ObservedTrait12, 0x20, descriptor.Trait12));
        return Variable(0, 0x88, fields.ToArray());
    }

    private sealed record ClassDescriptor(ushort Id, bool Marker6, bool Marker7, ushort? Trait9, uint Trait12);

    private enum PropertyField : byte {
        Magic = 1, DocumentId = 2, Version3 = 3, Version4 = 4, Version5 = 5,
        Version7 = 7, StreamLength = 8, FileName = 9, TrailerOffset = 10,
        CatalogOffset = 11, ObservedMarker12 = 12, Producer13 = 13, Producer14 = 14,
        Producer15 = 15, Producer16 = 16
    }

    private enum ClassField : byte {
        Id = 1, Version = 3, ObservedMarker6 = 6, ObservedMarker7 = 7,
        ObservedMarker8 = 8, ObservedTrait9 = 9, ObservedTrait12 = 12
    }
}
