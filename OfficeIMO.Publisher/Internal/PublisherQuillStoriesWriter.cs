using System.Text;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Rewrites ordinary source stories with preserved or allocated formatting pages, retaining unrelated native payloads.</summary>
internal static class PublisherQuillStoriesWriter {
    internal static byte[]? TryRewrite(byte[] source, IReadOnlyList<PublisherQuillStoryStyles> bindings,
        IReadOnlyDictionary<uint, PublisherTextStory> before, IReadOnlyDictionary<uint, PublisherTextStory> after,
        OfficeLegacyImportLimits limits, long maximumBytes, CancellationToken token) {
        var context = new PublisherParseContext(new PublisherReadOptions { Limits = limits }, token);
        PublisherQuillSourceLayout? layout = PublisherQuillSourceLayout.TryRead(source, context);
        if (layout == null) return null;
        PublisherBinaryData data = layout.Data; PublisherQuillChunk text = layout.Text, ids = layout.StoryIds, lengths = layout.StoryLengths;
        int count = bindings.Count;
        if (count == 0 || before.Count != count || after.Count != count || ids.Length != 8 + count * 4
            || data.U32(ids.Offset + 4) != count || lengths.Length != 22 + count * 8
            || data.U32(lengths.Offset) != count || data.U32(lengths.Offset + 4) != 8 || data.U32(lengths.Offset + 8) != 255) return null;
        var pc = new List<PublisherQuillFormattingRange>(); var pp = new List<PublisherQuillFormattingRange>();
        var replacement = new StringBuilder(); int originalStart = 0;
        PublisherQuillStoryStyles[] ordered = bindings.OrderBy(binding => binding.StoryOrdinal).ToArray();
        for (int ordinal = 0; ordinal < count; ordinal++) {
            token.ThrowIfCancellationRequested(); PublisherQuillStoryStyles binding = ordered[ordinal];
            if (binding.StoryOrdinal != ordinal || data.U32(ids.Offset + 8 + ordinal * 4) != binding.StoryId
                || !before.TryGetValue(binding.StoryId, out PublisherTextStory? prior) || !after.TryGetValue(binding.StoryId, out PublisherTextStory? current)
                || binding.NativeStart != originalStart || binding.NativeLength != prior.Text.Length || binding.NullCharacters != 0
                || binding.CarriageReturns != prior.Paragraphs.Count || prior.Paragraphs.Count == 0 || current.Paragraphs.Count == 0
                || !prior.HasFinalParagraphTerminator || !current.HasFinalParagraphTerminator || !current.Text.EndsWith("\n", StringComparison.Ordinal)
                || prior.Text.Count(c => c == '\n') != prior.Paragraphs.Count || current.Text.Count(c => c == '\n') != current.Paragraphs.Count
                || binding.TextByteOffset != text.Offset + originalStart * 2 || data.U32(lengths.Offset + 12 + ordinal * 4) != binding.NativeLength
                || !data.Utf16(binding.TextByteOffset, binding.NativeLength * 2).Equals(prior.Text.Replace('\n', '\r'), StringComparison.Ordinal)) return null;
            int outputStart = checked(text.Offset + replacement.Length * 2);
            if (!PublisherQuillParagraphWriter.TryPlan(data, context, binding, prior, current, outputStart, token, out var storyPc, out var storyPp)) return null;
            pc.AddRange(storyPc); pp.AddRange(storyPp);
            replacement.Append(current.Text.Replace('\n', '\r')); originalStart = checked(originalStart + binding.NativeLength);
        }
        if (originalStart * 2 != text.Length) return null;
        var characterPages = Pack(pc, layout.CharacterPages.Length, token);
        var paragraphPages = Pack(pp, layout.ParagraphPages.Length, token);
        byte[] encoded = Encoding.Unicode.GetBytes(replacement.ToString());
        if (characterPages == null || paragraphPages == null) {
            int capacity = PublisherQuillSourceLayout.MaximumDescriptors - layout.Chunks.Count + layout.CharacterPages.Length + layout.ParagraphPages.Length;
            characterPages = PublisherQuillPageAllocator.PackMinimum(pc, capacity, token);
            paragraphPages = PublisherQuillPageAllocator.PackMinimum(pp, capacity, token);
            return characterPages == null || paragraphPages == null ? null : PublisherQuillPageAllocator.TryWrite(layout, encoded,
                characterPages, paragraphPages, ordered.Select(binding => after[binding.StoryId].Text.Length).ToArray(), maximumBytes, token);
        }
        int shift = checked((Math.Max(0, encoded.Length - text.Length) + 511) / 512 * 512);
        int outputLength = checked(source.Length + shift);
        if (outputLength > maximumBytes) throw new InvalidDataException("Publisher native output byte limit exceeded.");
        token.ThrowIfCancellationRequested(); var output = new byte[outputLength];
        Buffer.BlockCopy(source, 0, output, 0, text.Offset);
        Buffer.BlockCopy(encoded, 0, output, text.Offset, encoded.Length);
        Buffer.BlockCopy(source, text.End, output, text.End + shift, source.Length - text.End);
        Write(20, (uint)outputLength);
        for (int i = 0; i < layout.Chunks.Count; i++) {
            token.ThrowIfCancellationRequested(); PublisherQuillChunk chunk = layout.Chunks[i];
            Write(32 + i * 24 + 16, (uint)(chunk == text ? text.Offset : chunk.Offset + shift));
            if (chunk == text) Write(32 + i * 24 + 20, (uint)encoded.Length);
        }
        for (int i = 0; i < count; i++) Write(lengths.Offset + shift + 12 + i * 4, (uint)after[ordered[i].StoryId].Text.Length);
        PublishPages(layout.CharacterPages, characterPages, layout.Find("BTEC"));
        PublishPages(layout.ParagraphPages, paragraphPages, layout.Find("BTEP"));
        token.ThrowIfCancellationRequested(); return output;

        void PublishPages(PublisherQuillChunk[] originals, IReadOnlyList<byte[]> pages, PublisherQuillChunk index) {
            for (int i = 0; i < originals.Length; i++) {
                token.ThrowIfCancellationRequested(); byte[] page = pages[i];
                Buffer.BlockCopy(page, 0, output, originals[i].Offset + shift, page.Length);
                if (i != 0) {
                    byte[] prior = pages[i - 1]; int priorCount = BitConverter.ToUInt16(prior, 0);
                    Write(originals[i].Offset + shift + 4, BitConverter.ToUInt32(prior, 8 + (priorCount - 1) * 4));
                }
                int ranges = BitConverter.ToUInt16(page, 0);
                Write(index.Offset + shift + 16 + i * 4, BitConverter.ToUInt32(page, 8 + (ranges - 1) * 4));
                Write(index.Offset + shift + 16 + originals.Length * 4 + i * 4, (uint)(originals[i].Offset + shift));
            }
        }
        void Write(int offset, uint value) { for (int i = 0; i < 4; i++) output[offset + i] = (byte)(value >> (i * 8)); }
    }

    private static IReadOnlyList<byte[]>? Pack(IReadOnlyList<PublisherQuillFormattingRange> ranges, int pageCount, CancellationToken token) {
        if (ranges.Count < pageCount) return null;
        var pages = new List<byte[]>(); int cursor = 0;
        for (int p = 0; p < pageCount; p++) {
            var current = new List<PublisherQuillFormattingRange>(); byte[]? accepted = null;
            int remainingPages = pageCount - p - 1;
            while (cursor < ranges.Count - remainingPages) {
                token.ThrowIfCancellationRequested(); current.Add(ranges[cursor]);
                byte[]? candidate = PublisherQuillFormattingPage.TryWriteShared(current);
                if (candidate == null) break;
                accepted = candidate; cursor++;
            }
            if (accepted == null) return null;
            pages.Add(accepted);
        }
        return cursor == ranges.Count ? pages : null;
    }
}
