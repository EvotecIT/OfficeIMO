namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillParagraphWriter {
    private static bool CanClearStory(PublisherTextStory before, PublisherTextStory after) => before.Paragraphs.Count > 0
        && before.Paragraphs.All(p => p.Label == null) && after.Paragraphs.Count == 1 && after.Paragraphs[0].Label == null
        && after.Paragraphs[0].Runs.Count == 0
        && PublisherTextComparison.ParagraphFormattingEqual(before.Paragraphs[before.Paragraphs.Count - 1], after.Paragraphs[0]);

    /// <summary>Native clearing retains the final complete paragraph definition and final CR, with no source body to invent.</summary>
    private static bool TryPlanClear(PublisherBinaryData data, PublisherParseContext context, PublisherQuillStoryStyles binding,
        PublisherTextStory before, int textOffset, CancellationToken token,
        List<PublisherQuillFormattingRange> pc, List<PublisherQuillFormattingRange> pp) {
        token.ThrowIfCancellationRequested();
        var paragraphs = binding.Uses.Where(u => u.Tag == "FDPP").ToArray();
        if (paragraphs.Length != before.Paragraphs.Count || binding.NativeLength < 1) return false;
        var last = paragraphs[paragraphs.Length - 1]; int final = checked(binding.NativeStart + binding.NativeLength - 1);
        if ((long)last.NativeStart + last.NativeLength != (long)final + 1) return false;
        var characters = binding.Uses.Where(u => u.Tag == "FDPC" && u.NativeStart <= final && (long)u.NativeStart + u.NativeLength > final).ToArray();
        if (characters.Length != 1) return false;
        context.Record(); int end = checked(textOffset + 2);
        pc.Add(new PublisherQuillFormattingRange(end, Copy(characters[0].DirectDefinition)));
        pp.Add(new PublisherQuillFormattingRange(end, Copy(last.DirectDefinition)));
        return true;
        byte[] Copy(PublisherQuillSourceSpan span) {
            var bytes = new byte[span.Length]; Buffer.BlockCopy(data.Bytes, span.Offset, bytes, 0, span.Length); return bytes;
        }
    }
}
