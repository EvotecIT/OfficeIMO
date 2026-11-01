using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Changes supported direct character fields, retaining every unrelated source token verbatim.</summary>
internal static partial class PublisherQuillCharacterStyleWriter {
    internal static byte[]? Rewrite(PublisherBinaryData source, PublisherQuillSourceSpan span,
        OfficeRichTextRun before, OfficeRichTextRun after, PublisherParseContext context, PublisherQuillWriteResources? resources = null) {
        bool size = before.FontSize != after.FontSize, bold = before.Bold != after.Bold;
        if (!size && !bold && !HasExtendedChanges(before, after) && !HasExplicitAttributeChange(before, after)) {
            var copy = new byte[span.Length];
            Buffer.BlockCopy(source.Bytes, span.Offset, copy, 0, span.Length);
            return copy;
        }
        var fields = new PublisherBlockReader(source, context).Chunk(span.Offset, span.Offset + span.Length);
        SortedDictionary<byte, byte[]>? extended = ExtendedFields(source, fields, before, after, context, resources);
        if (extended == null) return null;
        PublisherBlock[] sizes = fields.Where(field => field.Id == 0x0C).ToArray();
        PublisherBlock[] bolds = fields.Where(field => field.Id == 2).ToArray();
        double nativeSize = Math.Round(after.FontSize * 12700);
        if (size && (after.FontSize > 1000 || nativeSize <= 0 || after.FontSize != nativeSize / 12700D
                || sizes.Length > 1 || sizes.Any(field => field.Type != 0x22) || fields.Any(field => field.Id == 0x39))
            || bold && (bolds.Length > 1 || bolds.Any(field => field.Type is not (0x02 or 0x0A)) || fields.Any(field => field.Id == 0x37))) return null;
        if (bold) extended.Add(2, PublisherOplWriter.Boolean(2, after.Bold));
        if (size) extended.Add(0x0C, PublisherOplWriter.UInt32(0x0C, checked((uint)nativeSize)));
        return PublisherOplWriter.Rewrite(source, span.Offset, span.Length, fields, extended, context);
    }
}
