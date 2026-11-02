using OfficeIMO;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Binary;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTableBorderImagePreservationTests {
    [Fact]
    public void Border_topology_preserves_unrelated_embedded_image_loaded_with_a_higher_image_limit() {
        byte[] original = File.ReadAllBytes(PublisherNativeTests.Fixture("NativeComparison/TableDecoration/native-tabledecoration.pub"));
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out string? error), error);
        // Valid small PNG plus inert trailing bytes crosses the default encoded
        // image bound without introducing a large raster or a checked-in asset.
        byte[] small = PublisherPictureEditingFixture.Image(OfficeColor.Red);
        byte[] image = new byte[16 * 1024 * 1024 + 1];
        Buffer.BlockCopy(small, 0, image, 0, small.Length);
        byte[] entry = OfficeArtBlipStoreEntryWriter.CreateEmbedded(image, "image/png");
        byte[] store = new byte[8 + entry.Length];
        BitConverter.GetBytes((ushort)0x001F).CopyTo(store, 0);
        BitConverter.GetBytes((ushort)0xF001).CopyTo(store, 2);
        BitConverter.GetBytes((uint)entry.Length).CopyTo(store, 4);
        entry.CopyTo(store, 8);
        byte[] art = compound!.Streams["Escher/EscherStm"];
        byte[] withImage = new byte[art.Length + store.Length];
        Buffer.BlockCopy(art, 0, withImage, 0, 8);
        BitConverter.GetBytes(BitConverter.ToUInt32(art, 4) + (uint)store.Length).CopyTo(withImage, 4);
        store.CopyTo(withImage, 8);
        Buffer.BlockCopy(art, 8, withImage, 8 + store.Length, art.Length - 8);
        byte[] input = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = withImage });
        var options = new PublisherReadOptions { MaximumImageBytes = image.Length };
        PublisherDocument document = PublisherDocument.Load(input, options);
        Assert.Equal(image, Assert.Single(document.Images).GetBytes());
        PublisherTableBorderTopologyTests.Apply(document, "TableBorderTopology");
        byte[] saved = document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow });
        PublisherDocument reopened = PublisherDocument.Load(saved, options);
        Assert.Equal(image, Assert.Single(reopened.Images).GetBytes());
        Assert.Equal(PublisherTableBorderTopologyTests.Coverage(document.Pages[0].Tables[0]), PublisherTableBorderTopologyTests.Coverage(reopened.Pages[0].Tables[0]));
        PublisherNativeWritingTests.AssertPreservedStorage(input, saved, "Escher/EscherStm");
    }
}
