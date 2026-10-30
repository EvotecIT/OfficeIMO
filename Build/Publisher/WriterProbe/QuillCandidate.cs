using System.Text;
using OfficeIMO.Drawing;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Minimal Quill record hypothesis. Native required defaults remain subject to qualification.</summary>
internal static class QuillCandidate {
    private const int FirstPayloadOffset = 512;
    private const int MinimumStreamBytes = 1024;

    internal static byte[] Create(string? text, OfficeRichTextParagraph? paragraph) {
        var records = new List<QuillRecord>();
        if (text != null) {
            records.Add(new QuillRecord("TEXT", 0, Encoding.Unicode.GetBytes(text)));
            records.Add(new QuillRecord("SYID", 0, Bytes(writer => {
                writer.Write(NativePublicationGraph.TextStoryKey);
                writer.Write(1U); writer.Write(NativePublicationGraph.TextStoryKey);
            })));
            records.Add(new QuillRecord("STRS", 0, Bytes(writer => {
                writer.Write(1U); writer.Write(8U); writer.Write(0U); // Count, record size, clean flags hypothesis.
                writer.Write((uint)text.Length); writer.Write(0U); // Per-story length table plus final zero.
                writer.Write(Chunk()); // One self-inclusive, empty SDF record.
            })));
        }
        records.AddRange(QuillDefaultLibraries.Create());
        if (text != null) {
            // TEXT is first, so its absolute formatting endpoint follows its
            // actual serialized position even when the other libraries grow.
            int textEnd = FirstPayloadOffset + text.Length * 2;
            int firstPageOffset = Align(FirstPayloadOffset + records.Sum(record => record.Body.Length), QuillTextFormatting.PageBytes);
            IReadOnlyList<QuillFormattingRange> characterRanges = paragraph == null
                ? new[] { new QuillFormattingRange(textEnd, Chunk(Scalar(0x0C, 0x22, 18 * 12700))) }
                : QuillCharacterFormatting.FromParagraph(paragraph, FirstPayloadOffset);
            records.AddRange(QuillTextFormatting.CreateRecords(textEnd, firstPageOffset, characterRanges));
        }
        var positionedRecords = new List<(QuillRecord Record, int Offset)>();
        int payloadEnd = FirstPayloadOffset;
        foreach (var record in records) {
            payloadEnd = Align(payloadEnd, record.Alignment);
            positionedRecords.Add((record, payloadEnd));
            payloadEnd += record.Body.Length;
        }
        int logicalLength = Math.Max(MinimumStreamBytes, payloadEnd);
        byte[] header = Header(records.Count, logicalLength);
        int directoryEnd = header.Length + records.Count * 24;
        if (directoryEnd > FirstPayloadOffset) throw new InvalidDataException("The Quill directory exceeds its allocated envelope.");
        return Bytes(writer => {
            writer.Write(header);
            foreach (var (record, offset) in positionedRecords) {
                writer.Write((ushort)24); writer.Write(Encoding.ASCII.GetBytes(record.Tag));
                writer.Write(record.Id); writer.Write((ushort)1); writer.Write((ushort)0);
                writer.Write(Encoding.ASCII.GetBytes(record.Tag));
                writer.Write(offset); writer.Write(record.Body.Length);
            }
            writer.Write(new byte[FirstPayloadOffset - directoryEnd]);
            int offsetWritten = FirstPayloadOffset;
            foreach (var (record, offset) in positionedRecords) {
                writer.Write(new byte[offset - offsetWritten]);
                writer.Write(record.Body);
                offsetWritten = offset + record.Body.Length;
            }
            writer.Write(new byte[logicalLength - payloadEnd]);
        });
    }

    private static byte[] Header(int totalDescriptors, int logicalLength) => Bytes(writer => {
        writer.Write(Encoding.ASCII.GetBytes("CHNKINK "));
        // Native checks require version 4/7, byte +0x0F == 3 and a bounded,
        // seekable DWORD at +0x14 of at least 512. The remaining values follow
        // independently authored controls; their sufficiency is unqualified.
        writer.Write((ushort)4); writer.Write((ushort)7); writer.Write((ushort)totalDescriptors);
        writer.Write((byte)0); writer.Write((byte)3);
        writer.Write(FirstPayloadOffset); writer.Write(logicalLength);
        // The probe emits one directory; its count equals the total in +0x0C.
        writer.Write((ushort)(FirstPayloadOffset - 8)); writer.Write((ushort)totalDescriptors); writer.Write(uint.MaxValue);
    });

    private static int Align(int offset, int alignment) => checked((offset + alignment - 1) / alignment * alignment);
}

internal sealed record QuillRecord(string Tag, ushort Id, byte[] Body, int Alignment = 1);
