using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Rewrites ordinary source stories with preserved or allocated formatting pages, retaining unrelated native payloads.</summary>
internal static class PublisherQuillStoriesWriter {
    internal static byte[]? TryRewrite(byte[] source, IReadOnlyList<PublisherQuillStoryStyles> bindings,
        IReadOnlyDictionary<uint, PublisherTextStory> before, IReadOnlyDictionary<uint, PublisherTextStory> after,
        IReadOnlyDictionary<uint, PublisherTableStory> beforeTables, IReadOnlyDictionary<uint, PublisherTableStory> afterTables,
        IReadOnlyList<OfficeColor?> palette,
        OfficeLegacyImportLimits limits, long maximumBytes, CancellationToken token) {
        var context = new PublisherParseContext(new PublisherReadOptions { Limits = limits }, token);
        PublisherQuillSourceLayout? layout = PublisherQuillSourceLayout.TryRead(source, context);
        if (layout == null) return null;
        var resources = new PublisherQuillWriteResources(layout, context);
        PublisherBinaryData data = layout.Data; PublisherQuillChunk text = layout.Text, ids = layout.StoryIds, lengths = layout.StoryLengths;
        int count = bindings.Count;
        if (count == 0 || before.Count != count || after.Count != count || ids.Length != 8 + count * 4
            || data.U32(ids.Offset + 4) != count || lengths.Length != 22 + count * 8
            || data.U32(lengths.Offset) != count || data.U32(lengths.Offset + 4) != 8 || data.U32(lengths.Offset + 8) != 255) return null;
        if (!PublisherQuillTableTextWriter.TryPlan(layout, before, after, beforeTables, afterTables, token, out var tableChunks)) return null;
        // Empty styled runs must be compared with the actual inherited CR,
        // which can differ from the preceding visible body's style.
        PublisherQuillStyles? typingStyles = after.Values.Any(story => story.Paragraphs.Any(paragraph =>
            paragraph.Runs.Count == 1 && paragraph.Runs[0].Text.Length == 0))
            ? new PublisherQuillStyleReader(data, context, layout.Chunks, palette, text).Read() : null;
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
            if (!PublisherQuillParagraphWriter.TryPlan(data, context, binding, prior, current, outputStart, token, out var storyPc, out var storyPp, resources, typingStyles)) return null;
            pc.AddRange(storyPc); pp.AddRange(storyPp);
            replacement.Append(current.Text.Replace('\n', '\r')); originalStart = checked(originalStart + binding.NativeLength);
        }
        if (originalStart * 2 != text.Length) return null;
        var characterPages = Pack(pc, layout.CharacterPages.Length, token);
        var paragraphPages = Pack(pp, layout.ParagraphPages.Length, token);
        byte[] encoded = Encoding.Unicode.GetBytes(replacement.ToString());
        IReadOnlyDictionary<string, byte[]> replacements = resources.Replacements();
        if (characterPages == null || paragraphPages == null) {
            int capacity = PublisherQuillSourceLayout.MaximumDescriptors - layout.Chunks.Count + layout.CharacterPages.Length + layout.ParagraphPages.Length;
            characterPages = PublisherQuillPageAllocator.PackMinimum(pc, capacity, token);
            paragraphPages = PublisherQuillPageAllocator.PackMinimum(pp, capacity, token);
            return characterPages == null || paragraphPages == null ? null : PublisherQuillPageAllocator.TryWrite(layout, encoded,
                characterPages, paragraphPages, ordered.Select(binding => after[binding.StoryId].Text.Length).ToArray(), maximumBytes, token, replacements, tableChunks);
        }
        return PublisherQuillPageAllocator.TryWrite(layout, encoded, characterPages, paragraphPages,
            ordered.Select(binding => after[binding.StoryId].Text.Length).ToArray(), maximumBytes, token, replacements, tableChunks);
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
