using OfficeIMO.Publisher.Internal;
using System.Text;
using static OfficeIMO.Publisher.Internal.PublisherNativeCreationEncoding;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Minimal Quill record hypothesis. Native required defaults remain subject to qualification.</summary>
internal static class PublisherNativeCreationQuill {
    private const int FirstPayloadOffset = 512;
    private const int MinimumStreamBytes = 1024;

    internal static byte[] Create(PublisherNativeCreationGraph.Graph graph, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        IReadOnlyList<PublisherNativeCreationGraph.StoryBinding> stories = graph.Stories;
        var storyTexts = new string[stories.Count];
        StringBuilder? textBuilder = stories.Count == 0 ? null : new StringBuilder();
        uint maximumStoryKey = 0;
        for (int index = 0; index < stories.Count; index++) {
            context.Token.ThrowIfCancellationRequested();
            var story = stories[index];
            string storyText = StoryText(story.Source, context);
            storyTexts[index] = storyText;
            textBuilder!.Append(storyText);
            maximumStoryKey = Math.Max(maximumStoryKey, story.Key);
        }
        context.Token.ThrowIfCancellationRequested();
        string? text = textBuilder?.ToString();
        var records = new List<PublisherNativeCreationQuillRecord>();
        if (text != null) {
            context.Token.ThrowIfCancellationRequested();
            records.Add(new PublisherNativeCreationQuillRecord("TEXT", 0, Encoding.Unicode.GetBytes(text)));
            records.Add(new PublisherNativeCreationQuillRecord("SYID", 0, Bytes(writer => {
                writer.Write(maximumStoryKey);
                writer.Write((uint)stories.Count);
                foreach (var story in stories) {
                    context.Token.ThrowIfCancellationRequested();
                    writer.Write(story.Key);
                }
            })));
            records.Add(new PublisherNativeCreationQuillRecord("STRS", 0, Bytes(writer => {
                writer.Write((uint)stories.Count); writer.Write(8U); writer.Write(0U); // Count, record size, clean flags hypothesis.
                foreach (string storyText in storyTexts) {
                    context.Token.ThrowIfCancellationRequested();
                    writer.Write((uint)storyText.Length);
                }
                writer.Write(0U); // Final zero follows the per-story length table.
                foreach (var story in stories) {
                    context.Token.ThrowIfCancellationRequested();
                    writer.Write(PublisherOplWriter.Chunk()); // One self-inclusive, empty SDF per story.
                }
            })));
        }
        records.AddRange(PublisherNativeCreationQuillDefaults.Create(context));
        foreach (var story in stories) {
            context.Token.ThrowIfCancellationRequested();
            if (story.Source.CellEnds == null) continue;
            IReadOnlyList<int> ends = story.Source.CellEnds!;
            byte[] body = Bytes(writer => {
                writer.Write((uint)(ends.Count - 1)); writer.Write(0U); writer.Write(0xFF00U);
                for (int index = 0; index < ends.Count; index++) {
                    context.Token.ThrowIfCancellationRequested();
                    writer.Write((uint)(ends[index] - (index == ends.Count - 1 ? 0 : 1)));
                }
            });
            records.Add(new PublisherNativeCreationQuillRecord("TCD ", checked((ushort)(story.Key - 1)), body, DescriptorTag: "PLC "));
        }
        bool requiresLayoutCatalog = false;
        foreach (var frame in graph.Objects) {
            context.Token.ThrowIfCancellationRequested();
            if (frame.Source.FrameOrder != 0 || frame.Source.Table != null) { requiresLayoutCatalog = true; break; }
        }
        if (requiresLayoutCatalog)
            records.Add(new PublisherNativeCreationQuillRecord("MCLD", 0, PublisherNativeCreationLayoutCatalog.Create(graph, context)));
        if (text != null) {
            // TEXT is first, so its absolute formatting endpoint follows its
            // actual serialized position even when the other libraries grow.
            int textEnd = checked(FirstPayloadOffset + text.Length * 2);
            int recordBytes = 0;
            foreach (var record in records) {
                context.Token.ThrowIfCancellationRequested();
                recordBytes = checked(recordBytes + record.Body.Length);
            }
            int firstPageOffset = Align(checked(FirstPayloadOffset + recordBytes), PublisherQuillFormattingPage.PageBytes);
            var characterRanges = new List<PublisherQuillFormattingRange>();
            var paragraphRanges = new List<PublisherQuillFormattingRange>();
            int storyStart = FirstPayloadOffset;
            for (int storyIndex = 0; storyIndex < stories.Count; storyIndex++) {
                context.Token.ThrowIfCancellationRequested();
                var story = stories[storyIndex];
                int storyEnd = checked(storyStart + storyTexts[storyIndex].Length * 2);
                if (story.Source.Paragraphs == null) {
                    characterRanges.Add(new PublisherQuillFormattingRange(storyEnd, PublisherOplWriter.Chunk(PublisherOplWriter.UInt32(0x0C, 18 * 12700, 0x22))));
                    paragraphRanges.Add(new PublisherQuillFormattingRange(storyEnd, PublisherOplWriter.Chunk()));
                } else {
                    int paragraphStart = storyStart;
                    foreach (var paragraph in story.Source.Paragraphs) {
                        context.Token.ThrowIfCancellationRequested();
                        characterRanges.AddRange(PublisherNativeCreationCharacterFormatting.FromParagraph(paragraph, paragraphStart, context));
                        int paragraphCharacters = paragraph.Runs.Sum(run => {
                            context.Token.ThrowIfCancellationRequested();
                            return run.Text.Length;
                        });
                        paragraphStart = checked(paragraphStart + (paragraphCharacters + 1) * 2);
                        paragraphRanges.Add(new PublisherQuillFormattingRange(paragraphStart, PublisherNativeCreationParagraphFormatting.Create(paragraph, context)));
                    }
                }
                storyStart = storyEnd;
            }
            int formattingCapacity = PublisherQuillSourceLayout.MaximumDescriptors - records.Count - 2;
            records.AddRange(PublisherNativeCreationTextFormatting.CreateRecords(textEnd, firstPageOffset, characterRanges, paragraphRanges, formattingCapacity, context));
        }
        var positionedRecords = new List<(PublisherNativeCreationQuillRecord Record, int Offset)>();
        int payloadEnd = FirstPayloadOffset;
        foreach (var record in records) {
            context.Token.ThrowIfCancellationRequested();
            payloadEnd = Align(payloadEnd, record.Alignment);
            positionedRecords.Add((record, payloadEnd));
            payloadEnd = checked(payloadEnd + record.Body.Length);
        }
        int logicalLength = Math.Max(MinimumStreamBytes, payloadEnd);
        PublisherQuillDirectoryWriter directories = PublisherQuillDirectoryWriter.PlanFresh(records.Count, () => {
            context.Token.ThrowIfCancellationRequested();
            int offset = Align(logicalLength, PublisherQuillFormattingPage.PageBytes);
            logicalLength = checked(offset + PublisherQuillFormattingPage.PageBytes);
            return offset;
        }, context.Token);
        context.Token.ThrowIfCancellationRequested();
        var output = new byte[logicalLength]; Header().CopyTo(output, 0);
        var rows = new List<PublisherQuillFreshDescriptor>();
        foreach (var (record, offset) in positionedRecords) {
            context.Token.ThrowIfCancellationRequested();
            Buffer.BlockCopy(record.Body, 0, output, offset, record.Body.Length);
            rows.Add(new(record.Tag, record.Id, record.DescriptorTag ?? record.Tag, offset, record.Body.Length));
        }
        directories.PublishFresh(output, rows, context.Token);
        context.Token.ThrowIfCancellationRequested();
        return output;
    }

    private static string StoryText(PublisherNativeCreationStory story, PublisherParseContext context) {
        context.Token.ThrowIfCancellationRequested();
        if (story.Paragraphs == null) {
            string plain = story.Text;
            context.Token.ThrowIfCancellationRequested();
            return plain;
        }
        var builder = new StringBuilder();
        foreach (var paragraph in story.Paragraphs) {
            context.Token.ThrowIfCancellationRequested();
            foreach (var run in paragraph.Runs) {
                context.Token.ThrowIfCancellationRequested();
                builder.Append(run.Text);
            }
            builder.Append('\r');
        }
        context.Token.ThrowIfCancellationRequested();
        string result = builder.ToString();
        context.Token.ThrowIfCancellationRequested();
        return result;
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

internal sealed record PublisherNativeCreationQuillRecord(string Tag, ushort Id, byte[] Body, int Alignment = 1, string? DescriptorTag = null);
