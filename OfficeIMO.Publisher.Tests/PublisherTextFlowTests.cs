using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherTextFlowTests {
    [Theory]
    [InlineData(18U, new uint[] { 326, 327, 328 })]
    [InlineData(22U, new uint[] { 330, 329, 331 })]
    [InlineData(31U, new uint[] { 342, 340, 341 })]
    public void Native_story_links_define_continuation_order_independently_of_drawing_order(uint storyId, uint[] order) {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        PublisherTextFrame[] frames = document.Pages.SelectMany(page => page.TextFrames)
            .Where(frame => frame.StoryId == storyId).OrderBy(frame => frame.Order).ToArray();
        Assert.Equal(order, frames.Select(frame => frame.Id));
        Assert.Equal(0, frames[0].TextStart); Assert.Null(frames[0].PreviousFrameId); Assert.Null(frames[frames.Length - 1].NextFrameId);
        for (int i = 1; i < frames.Length; i++) {
            Assert.Equal(frames[i - 1].Id, frames[i].PreviousFrameId); Assert.Equal(frames[i].Id, frames[i - 1].NextFrameId);
            Assert.Equal(frames[i - 1].TextStart + frames[i - 1].TextLength, frames[i].TextStart);
        }
        Assert.Contains(frames.Skip(1), frame => frame.TextLength > 0);
        Assert.DoesNotContain(document.ReadReport.FidelityDiagnostics, item => item.Code == "PUB_LINKED_TEXT_FRAME_OMITTED");
    }

    [Fact]
    public void Recovered_frame_ranges_neither_duplicate_nor_skip_source_text() {
        PublisherDocument document = PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"));
        PublisherTextStory story = document.TextStories.Single(item => item.Id == 18);
        PublisherTextFrame[] frames = document.Pages.SelectMany(page => page.TextFrames).Where(frame => frame.StoryId == story.Id)
            .OrderBy(frame => frame.Order).ToArray();
        string recovered = string.Concat(frames.Select(frame => story.Text.Substring(frame.TextStart!.Value, frame.TextLength!.Value)));
        Assert.StartsWith(recovered, story.Text);
        Assert.Equal(frames.Sum(frame => frame.TextLength), recovered.Length);
        foreach (PublisherTextFrame frame in frames.Where(frame => frame.TextLength > 0)) {
            string painted = string.Concat(PublisherNativeTests.Elements(document.Pages.Single(page => page.Id == frame.PageId).Drawing)
                .OfType<OfficeDrawingRichText>().Where(text => text.SourceElementIds?.Contains("publisher-object-" + frame.Id) == true)
                .Select(PublisherNativeTests.Text));
            string assigned = story.Text.Substring(frame.TextStart!.Value, frame.TextLength!.Value);
            Assert.Equal(assigned.Replace("\n", string.Empty), painted.Replace("\n", string.Empty));
        }
        Assert.Equal(recovered.Length < story.Text.Length, frames[frames.Length - 1].HasOverflow);
    }

    [Fact]
    public void Native_multi_column_properties_produce_separate_ordered_content_regions() {
        byte[] input = PublisherInputContractTests.Mutate("Escher/EscherStm", bytes => {
            Assert.True(ChangeColumnCount(bytes, 0, bytes.Length, 293, 2));
        });
        PublisherDocument document = PublisherDocument.Load(input);
        PublisherTextFrame frame = document.Pages[0].TextFrames.Single(item => item.Id == 293);
        Assert.Equal(2, frame.ColumnCount);
        OfficeDrawingRichText[] columns = PublisherNativeTests.Elements(document.Pages[0].Drawing).OfType<OfficeDrawingRichText>()
            .Where(item => item.SourceElementIds?.Contains("publisher-object-293") == true).ToArray();
        Assert.Equal(2, columns.Length); Assert.True(columns[0].X < columns[1].X);
        Assert.Equal(columns[0].Width, columns[1].Width); Assert.Equal(columns[0].Y, columns[1].Y);
        Assert.StartsWith("0123456789", PublisherNativeTests.Text(columns[0]));
        PublisherTextStory story = document.TextStories.Single(item => item.Id == frame.StoryId);
        Assert.Equal(story.Text.Substring(0, frame.TextLength!.Value).Replace("\n", string.Empty),
            string.Concat(columns.Select(PublisherNativeTests.Text)).Replace("\n", string.Empty));
    }

    [Theory]
    [InlineData(329U, (byte)0x36, 329U)] // Self-link.
    [InlineData(330U, (byte)0x37, 340U)] // Different story.
    [InlineData(329U, (byte)0x28, 2U)] // Wrong chain ordinal.
    public void Corrupt_native_frame_links_are_rejected_without_guessed_placement(uint frameId, byte field, uint value) {
        byte[] input = PublisherInputContractTests.Mutate("Contents", bytes => {
            int offset = NativeFrameField(bytes, frameId, field);
            PublisherInputContractTests.WriteUInt32(bytes, offset, value);
        }, "SampleNewsletter.pub");
        Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(input));
    }

    [Fact]
    public void Repeated_story_measurement_obeys_the_configured_work_ceiling() {
        var options = new PublisherReadOptions { Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 10_000 } };
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => PublisherDocument.Load(PublisherNativeTests.Fixture("SampleNewsletter.pub"), options));
        Assert.Contains("text layout character work limit", error.Message);
    }

    internal static int NativeFrameField(byte[] bytes, uint objectId, byte id) {
        int trailer = checked((int)BitConverter.ToUInt32(bytes, 0x1A));
        int directory = Blocks(bytes, trailer + 4, trailer + checked((int)BitConverter.ToUInt32(bytes, trailer)))
            .Single(block => block.Type == 0x90).Offset;
        var slots = Blocks(bytes, directory + 4, directory + checked((int)BitConverter.ToUInt32(bytes, directory))).ToArray();
        var slot = slots[checked((int)objectId)];
        int shape = checked((int)BitConverter.ToUInt32(bytes,
            Blocks(bytes, slot.Offset + 4, slot.Offset + checked((int)BitConverter.ToUInt32(bytes, slot.Offset))).Single(block => block.Id == 4).Offset));
        return Blocks(bytes, shape + 4, shape + checked((int)BitConverter.ToUInt32(bytes, shape))).Single(block => block.Id == id).Offset;
    }
    private static IEnumerable<(byte Id, byte Type, int Offset)> Blocks(byte[] bytes, int start, int end) {
        for (int cursor = start; cursor < end;) {
            byte type = bytes[cursor + 1]; int offset = cursor + 2;
            int length = type switch {
                0x00 or 0x02 or 0x05 or 0x08 or 0x0A or 0x78 => 0,
                0x07 or 0x10 or 0x12 or 0x18 or 0x1A => 2,
                0x20 or 0x22 or 0x58 or 0x68 or 0x70 or 0xB8 => 4,
                0x28 => 8, 0x38 => 16, 0x48 => 24,
                _ => checked((int)BitConverter.ToUInt32(bytes, offset))
            };
            yield return (bytes[cursor], type, offset); cursor = offset + length;
        }
    }
    private static bool ChangeColumnCount(byte[] bytes, int start, int end, uint objectId, uint columns) {
        for (int offset = start; offset < end;) {
            ushort initial = BitConverter.ToUInt16(bytes, offset), kind = BitConverter.ToUInt16(bytes, offset + 2);
            int content = offset + 8, boundary = content + checked((int)BitConverter.ToUInt32(bytes, offset + 4));
            if (kind == 0xF004) {
                int setting = -1, anchor = -1; bool matches = false;
                for (int child = content; child < boundary;) {
                    ushort type = BitConverter.ToUInt16(bytes, child + 2);
                    int length = checked((int)BitConverter.ToUInt32(bytes, child + 4));
                    if (type == 0xF011 && length == 10) matches = BitConverter.ToUInt32(bytes, child + 14) == objectId;
                    if (type == 0xF010) anchor = child + 8;
                    if (type == 0xF122) {
                        int count = BitConverter.ToUInt16(bytes, child) >> 4;
                        for (int property = 0; property < count; property++) {
                            int position = child + 8 + property * 6;
                            if ((BitConverter.ToUInt16(bytes, position) & 0x3FFF) == 0x8D) setting = position;
                        }
                    }
                    child += 8 + length;
                }
                if (matches && setting >= 0) {
                    bytes[setting] = 0x8C; bytes[setting + 1] = 0;
                    PublisherInputContractTests.WriteUInt32(bytes, setting + 2, columns);
                    if (anchor >= 0) PublisherInputContractTests.WriteUInt32(bytes, anchor + 24,
                        unchecked((uint)(BitConverter.ToInt32(bytes, anchor + 12) + 254_000)));
                    return true;
                }
            } else if ((initial & 15) == 15 && ChangeColumnCount(bytes, content, boundary, objectId, columns)) return true;
            offset = boundary + (kind is 0xF000 or 0xF002 && boundary < end ? 4 : 0);
        }
        return false;
    }
}
