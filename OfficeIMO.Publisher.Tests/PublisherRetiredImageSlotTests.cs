using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing.Binary;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherRetiredImageSlotTests {
    private const string Fixture = "NativeComparison/ImageRetiredSlot/native-imageretiredslot.pub";

    [Fact]
    public void Retired_slot_inspection_still_accounts_payload_work_before_discarding_stale_image_data() {
        byte[] original = File.ReadAllBytes(PublisherNativeTests.Fixture(Fixture));
        Assert.True(OfficeCompoundFileReader.TryRead(original, out OfficeCompoundFile? compound, out string? error), error);
        byte[] source = compound!.Streams["Escher/EscherStm"];
        var data = new PublisherBinaryData(source, "Escher/EscherStm");
        var context = new PublisherParseContext(new PublisherReadOptions(), default);
        var reader = new PublisherEscherReader(data, null, context);
        PublisherEscherRecord group = reader.Records(0, source.Length).Single(record => record.Kind == 0xF000);
        PublisherEscherRecord store = reader.Records(group.Offset, group.End).Single(record => record.Kind == 0xF001);
        PublisherEscherRecord first = reader.Records(store.Offset, store.End)[0];
        const int copies = 1000;
        var replacement = new byte[8 + store.Length + copies * (first.Length + 8)];
        Buffer.BlockCopy(source, store.Offset - 8, replacement, 0, store.Length + 8);
        BitConverter.GetBytes((ushort)((((store.Initial >> 4) + copies) << 4) | 15)).CopyTo(replacement, 0);
        BitConverter.GetBytes((uint)(replacement.Length - 8)).CopyTo(replacement, 4);
        for (int index = 0; index < copies; index++) Buffer.BlockCopy(source, first.Offset - 8, replacement,
            8 + store.Length + index * (first.Length + 8), first.Length + 8);
        byte[]? artwork = PublisherEscherSourceWriter.Rewrite(source,
            new Dictionary<int, IReadOnlyDictionary<ushort, OfficeArtProperty>>(),
            new PublisherNativeStreamEdit(store.Offset - 8, store.Length + 8, replacement), context, 64 * 1024 * 1024);
        Assert.NotNull(artwork);
        byte[] input = OfficeCompoundFileWriter.Rewrite(compound, new Dictionary<string, byte[]> { ["Escher/EscherStm"] = artwork! });
        Assert.True(input.Length < copies * 200, "The encoded input must fit below the repeated stale-payload work.");
        var exception = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input,
            new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxInputBytes = input.Length } }));
        Assert.Contains("image processing byte limit", exception.Message);
    }

    [Fact]
    public void Native_resaved_retired_image_slot_stays_reserved_and_inert_while_active_images_can_be_edited_and_saved_again() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture(Fixture));
        Assert.DoesNotContain(document.Images, image => image.Id == 1);
        Assert.Equal(new[] { 2, 3, 4 }, document.Images.Select(image => image.Id));
        Assert.Contains(document.SourceReadReport.FidelityDiagnostics, diagnostic => diagnostic.Code == "PUB_IMAGE_RETIRED_SLOT"
            && diagnostic.LossKind == OfficeConversionLossKind.None && diagnostic.Location == "Escher/image/1");
        PublisherPicture picture = document.Pages[0].Pictures[0];
        byte[] replacement = PublisherPictureEditingFixture.Image(OfficeColor.Red, 16, 24);
        document.ReplacePictureImage(picture.Id, replacement);
        PublisherPicture current = document.Pages[0].Pictures[0];
        Assert.Equal(5, current.ImageId);
        PublisherDocument reopened = PublisherDocument.Load(document.ToPubBytes(new PublisherSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        PublisherPicture saved = reopened.Pages[0].Pictures[0];
        Assert.Equal(picture.Id, saved.Id); Assert.Equal(5, saved.ImageId);
        Assert.Equal(replacement, reopened.Images.Single(image => image.Id == saved.ImageId).GetBytes());
        Assert.DoesNotContain(reopened.Images, image => image.Id == 1);
        Assert.Equal(document.TextStories.Select(story => story.Text), reopened.TextStories.Select(story => story.Text));
    }

    [Theory]
    [InlineData("picture")]
    [InlineData("counter")]
    [InlineData("unbound-picture")]
    public void Retired_slot_recovery_does_not_hide_a_truncated_image_with_a_native_reference_or_positive_counter(string reference) {
        byte[] input;
        if (reference == "picture") input = PublisherDrawingFixture.Mutate(new Dictionary<ushort, uint> { [0x4104] = 1 },
            1, fixture: Fixture, objectId: 298);
        else input = PublisherInputContractTests.Mutate("Escher/EscherStm", bytes => {
            var data = new PublisherBinaryData(bytes, "Escher/EscherStm");
            var reader = new PublisherEscherReader(data, null, new PublisherParseContext(new PublisherReadOptions(), default));
            if (reference == "counter") {
                PublisherEscherRecord group = reader.Records(0, bytes.Length).Single(record => record.Kind == 0xF000);
                PublisherEscherRecord store = reader.Records(group.Offset, group.End).Single(record => record.Kind == 0xF001);
                PublisherEscherRecord first = reader.Records(store.Offset, store.End)[0];
                BitConverter.GetBytes(1U).CopyTo(bytes, first.Offset + 24);
            } else {
                IEnumerable<PublisherEscherRecord> Shapes(int start, int end) {
                    foreach (PublisherEscherRecord record in reader.Records(start, end)) {
                        if (record.Kind == 0xF004) yield return record;
                        else if ((record.Initial & 15) == 15 && record.Kind != 0xF001)
                            foreach (PublisherEscherRecord child in Shapes(record.Offset, record.End)) yield return child;
                    }
                }
                bool changed = false;
                foreach (PublisherEscherRecord shape in Shapes(0, bytes.Length)) {
                    var children = reader.Records(shape.Offset, shape.End);
                    PublisherEscherRecord? client = children.SingleOrDefault(record => record.Kind == 0xF011);
                    if (client == null || client.Length != 10 || data.U32(client.Offset + 6) != 298) continue;
                    BitConverter.GetBytes(0U).CopyTo(bytes, client.Offset + 6);
                    foreach (var table in children.Where(record => record.Kind is 0xF00B or 0xF122))
                        for (int index = 0; index < table.Initial >> 4; index++)
                            if (data.U16(table.Offset + index * 6) == 0x4104) {
                                BitConverter.GetBytes(1U).CopyTo(bytes, table.Offset + index * 6 + 2); changed = true;
                            }
                }
                Assert.True(changed);
            }
        }, Fixture);
        var error = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input));
        Assert.Contains("Truncated Publisher image payload", error.Message);
    }
}
