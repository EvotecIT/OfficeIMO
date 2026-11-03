using OfficeIMO.Publisher.Internal;
using System.Text;
using static OfficeIMO.Publisher.WriterProbe.NativeEncoding;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Minimal Quill record hypothesis. Native required defaults remain subject to qualification.</summary>
internal static class QuillCandidate {
    private const int FirstPayloadOffset = 512;
    private const int MinimumStreamBytes = 1024;

    internal static byte[] Create(NativePublicationGraph.Graph graph) {
        IReadOnlyList<NativePublicationGraph.StoryBinding> stories = graph.Stories;
        string? text = stories.Count == 0 ? null : string.Concat(stories.Select(story => story.Source.Text));
        var records = new List<QuillRecord>();
        if (text != null) {
            records.Add(new QuillRecord("TEXT", 0, Encoding.Unicode.GetBytes(text)));
            records.Add(new QuillRecord("SYID", 0, Bytes(writer => {
                writer.Write(stories.Max(story => story.Key));
                writer.Write((uint)stories.Count);
                foreach (var story in stories) writer.Write(story.Key);
            })));
            records.Add(new QuillRecord("STRS", 0, Bytes(writer => {
                writer.Write((uint)stories.Count); writer.Write(8U); writer.Write(0U); // Count, record size, clean flags hypothesis.
                foreach (var story in stories) writer.Write((uint)story.Source.Text.Length);
                writer.Write(0U); // Final zero follows the per-story length table.
                foreach (var story in stories) writer.Write(Chunk()); // One self-inclusive, empty SDF per story.
            })));
        }
        records.AddRange(QuillDefaultLibraries.Create());
        foreach (var story in stories.Where(story => story.Source.CellEnds != null)) {
            IReadOnlyList<int> ends = story.Source.CellEnds!;
            byte[] body = Bytes(writer => {
                writer.Write((uint)(ends.Count - 1)); writer.Write(0U); writer.Write(0xFF00U);
                for (int index = 0; index < ends.Count; index++)
                    writer.Write((uint)(ends[index] - (index == ends.Count - 1 ? 0 : 1)));
            });
            records.Add(new QuillRecord("TCD ", checked((ushort)(story.Key - 1)), body, DescriptorTag: "PLC "));
        }
        if (graph.Objects.Any(frame => frame.Source.FrameOrder != 0 || frame.Source.Table != null))
            records.Add(new QuillRecord("MCLD", 0, QuillLayoutCatalog.Create(graph)));
        if (text != null) {
            // TEXT is first, so its absolute formatting endpoint follows its
            // actual serialized position even when the other libraries grow.
            int textEnd = checked(FirstPayloadOffset + text.Length * 2);
            int firstPageOffset = Align(checked(FirstPayloadOffset + records.Sum(record => record.Body.Length)), PublisherQuillFormattingPage.PageBytes);
            var characterRanges = new List<PublisherQuillFormattingRange>();
            var paragraphRanges = new List<PublisherQuillFormattingRange>();
            int storyStart = FirstPayloadOffset;
            foreach (var story in stories) {
                int storyEnd = checked(storyStart + story.Source.Text.Length * 2);
                if (story.Source.Paragraphs == null) {
                    characterRanges.Add(new PublisherQuillFormattingRange(storyEnd, Chunk(Scalar(0x0C, 0x22, 18 * 12700))));
                    paragraphRanges.Add(new PublisherQuillFormattingRange(storyEnd, Chunk()));
                } else {
                    int paragraphStart = storyStart;
                    foreach (var paragraph in story.Source.Paragraphs) {
                        characterRanges.AddRange(QuillCharacterFormatting.FromParagraph(paragraph, paragraphStart));
                        paragraphStart = checked(paragraphStart + (paragraph.Runs.Sum(run => run.Text.Length) + 1) * 2);
                        paragraphRanges.Add(new PublisherQuillFormattingRange(paragraphStart, QuillParagraphFormatting.Create(paragraph)));
                    }
                }
                storyStart = storyEnd;
            }
            int formattingCapacity = PublisherQuillSourceLayout.MaximumDescriptors - records.Count - 2;
            records.AddRange(QuillTextFormatting.CreateRecords(textEnd, firstPageOffset, characterRanges, paragraphRanges, formattingCapacity));
        }
        var positionedRecords = new List<(QuillRecord Record, int Offset)>();
        int payloadEnd = FirstPayloadOffset;
        foreach (var record in records) {
            payloadEnd = Align(payloadEnd, record.Alignment);
            positionedRecords.Add((record, payloadEnd));
            payloadEnd = checked(payloadEnd + record.Body.Length);
        }
        int logicalLength = Math.Max(MinimumStreamBytes, payloadEnd);
        PublisherQuillDirectoryWriter directories = PublisherQuillDirectoryWriter.PlanFresh(records.Count, () => {
            int offset = Align(logicalLength, PublisherQuillFormattingPage.PageBytes);
            logicalLength = checked(offset + PublisherQuillFormattingPage.PageBytes);
            return offset;
        }, CancellationToken.None);
        var output = new byte[logicalLength]; Header().CopyTo(output, 0);
        var rows = new List<PublisherQuillFreshDescriptor>();
        foreach (var (record, offset) in positionedRecords) {
            Buffer.BlockCopy(record.Body, 0, output, offset, record.Body.Length);
            rows.Add(new(record.Tag, record.Id, record.DescriptorTag ?? record.Tag, offset, record.Body.Length));
        }
        directories.PublishFresh(output, rows, CancellationToken.None);
        return output;
    }

    private static byte[] Header() => Bytes(writer => {
        writer.Write(Encoding.ASCII.GetBytes("CHNKINK "));
        // Shared directory publication owns counts, lengths and page links.
        // This prefix retains the independently authored version/reserved fields.
        writer.Write((ushort)4); writer.Write((ushort)7); writer.Write((ushort)0);
        writer.Write((byte)0); writer.Write((byte)3);
    });

    private static int Align(int offset, int alignment) => checked((offset + alignment - 1) / alignment * alignment);
}

internal sealed record QuillRecord(string Tag, ushort Id, byte[] Body, int Alignment = 1, string? DescriptorTag = null);
