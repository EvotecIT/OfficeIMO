using System.Text;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherQuillFreshDirectoryTests {
    [Theory]
    [InlineData(20)]
    [InlineData(21)]
    [InlineData(41)]
    public void Fresh_directory_chain_retains_text_position_payload_and_complete_descriptor_metadata(int count) {
        var rows = Enumerable.Range(0, count).Select(index => new PublisherQuillFreshDescriptor(
            index == 0 ? "TEXT" : "TCD ", (ushort)index, index == 0 ? "TEXT" : "PLC ", 512 + index * 2, 2)).ToArray();
        int end = 1024;
        PublisherQuillDirectoryWriter plan = PublisherQuillDirectoryWriter.PlanFresh(count, () => {
            int offset = end; end += 512; return offset;
        }, CancellationToken.None);
        byte[] bytes = new byte[end];
        Encoding.ASCII.GetBytes("CHNKINK ").CopyTo(bytes, 0);
        byte[] payload = Enumerable.Range(0, count * 2).Select(index => (byte)(index + 1)).ToArray();
        payload.CopyTo(bytes, 512);
        plan.PublishFresh(bytes, rows, CancellationToken.None);
        Assert.Equal(payload, bytes.Skip(512).Take(payload.Length));
        Assert.Equal(count, BitConverter.ToUInt16(bytes, 12));
        int directories = (count + 19) / 20;
        Assert.Equal((uint)(directories * 512), BitConverter.ToUInt32(bytes, 16));
        Assert.Equal((uint)bytes.Length, BitConverter.ToUInt32(bytes, 20));
        int remaining = count, ordinal = 0;
        for (int page = 0; page < directories; page++) {
            int offset = page == 0 ? 24 : 1024 + (page - 1) * 512;
            int assigned = Math.Min(20, remaining);
            Assert.Equal(504, BitConverter.ToUInt16(bytes, offset));
            Assert.Equal(assigned, BitConverter.ToUInt16(bytes, offset + 2));
            Assert.Equal(page == directories - 1 ? uint.MaxValue : (uint)(1024 + page * 512), BitConverter.ToUInt32(bytes, offset + 4));
            for (int index = 0; index < assigned; index++) {
                int descriptor = offset + 8 + index * 24;
                Assert.Equal(24, BitConverter.ToUInt16(bytes, descriptor));
                Assert.Equal(1, BitConverter.ToUInt16(bytes, descriptor + 8));
                Assert.Equal(0, BitConverter.ToUInt16(bytes, descriptor + 10));
                Assert.Equal(rows[ordinal++].DescriptorTag, Encoding.ASCII.GetString(bytes, descriptor + 12, 4));
            }
            remaining -= assigned;
        }
        Assert.Equal(count, ordinal);
        var context = new PublisherParseContext(new PublisherReadOptions(), CancellationToken.None);
        var actual = PublisherQuillDirectory.Read(new PublisherBinaryData(bytes, "Quill/QuillSub/CONTENTS"), context);
        Assert.Equal(rows.Select(row => (row.Tag, (int)row.Id, row.Offset, row.Length)),
            actual.Select(row => (row.Tag, row.Id, row.Offset, row.Length)));
        Assert.Equal(512, actual[0].Offset);
    }
}
