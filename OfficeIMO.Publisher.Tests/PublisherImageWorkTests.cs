using OfficeIMO.Core.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherImageWorkTests {
    [Fact]
    public void Repeated_unrecoverable_delayed_images_consume_a_bounded_processing_budget() {
        byte[] input = WithRepeatedDelayedImages(256, 1024);
        var error = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input, new PublisherReadOptions {
            Limits = new OfficeLegacyImportLimits { MaxInputBytes = input.Length }
        }));
        Assert.Contains("image processing byte limit", error.Message);
    }

    [Fact]
    public void Unrecoverable_image_store_entries_still_obey_the_item_limit() {
        byte[] input = WithRepeatedDelayedImages(256, 32);
        var error = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input, new PublisherReadOptions {
            Limits = new OfficeLegacyImportLimits { MaxItems = 128 }
        }));
        Assert.Contains("image store entry limit", error.Message);
    }

    internal static byte[] WithRepeatedDelayedImages(int count, int payloadBytes) {
        Assert.True(OfficeCompoundFileReader.TryRead(File.ReadAllBytes(PublisherNativeTests.Fixture("Simple.pub")),
            out OfficeCompoundFile? source, out string? error), error);
        byte[] original = source!.Streams["Escher/EscherStm"];
        var store = new byte[8 + count * 44];
        WriteUInt16(store, 0, (ushort)((Math.Min(count, 4095) << 4) | 15));
        WriteUInt16(store, 2, 0xF001);
        PublisherInputContractTests.WriteUInt32(store, 4, (uint)(store.Length - 8));
        for (int i = 0; i < count; i++) {
            int offset = 8 + i * 44;
            WriteUInt16(store, offset, 0x42);
            WriteUInt16(store, offset + 2, 0xF007);
            PublisherInputContractTests.WriteUInt32(store, offset + 4, 36);
            store[offset + 8] = store[offset + 9] = 4; // Unsupported PICT remains metadata, never an imported image.
            PublisherInputContractTests.WriteUInt32(store, offset + 28, (uint)(payloadBytes + 8));
            PublisherInputContractTests.WriteUInt32(store, offset + 32, 1);
        }
        var escher = new byte[original.Length + store.Length];
        Buffer.BlockCopy(original, 0, escher, 0, 8);
        PublisherInputContractTests.WriteUInt32(escher, 4, BitConverter.ToUInt32(original, 4) + (uint)store.Length);
        Buffer.BlockCopy(store, 0, escher, 8, store.Length);
        Buffer.BlockCopy(original, 8, escher, 8 + store.Length, original.Length - 8);
        var delayed = new byte[payloadBytes + 8];
        WriteUInt16(delayed, 0, 0x5420);
        WriteUInt16(delayed, 2, 0xF01C);
        PublisherInputContractTests.WriteUInt32(delayed, 4, (uint)payloadBytes);
        return OfficeCompoundFileWriter.Rewrite(source, new Dictionary<string, byte[]> {
            ["Escher/EscherStm"] = escher, ["Escher/EscherDelayStm"] = delayed
        });
    }
    private static void WriteUInt16(byte[] bytes, int offset, ushort value) {
        bytes[offset] = (byte)value; bytes[offset + 1] = (byte)(value >> 8);
    }
}
