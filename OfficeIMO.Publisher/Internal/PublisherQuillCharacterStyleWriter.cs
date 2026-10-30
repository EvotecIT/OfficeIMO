using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Changes supported direct character fields, retaining every unrelated source token verbatim.</summary>
internal static partial class PublisherQuillCharacterStyleWriter {
    internal static byte[]? Rewrite(PublisherBinaryData source, PublisherQuillSourceSpan span,
        OfficeRichTextRun before, OfficeRichTextRun after, PublisherParseContext context, PublisherQuillWriteResources? resources = null) {
        bool size = before.FontSize != after.FontSize, bold = before.Bold != after.Bold;
        if (!size && !bold && !HasExtendedChanges(before, after)) {
            var copy = new byte[span.Length];
            Buffer.BlockCopy(source.Bytes, span.Offset, copy, 0, span.Length);
            return copy;
        }
        var fields = new PublisherBlockReader(source, context).Chunk(span.Offset, span.Offset + span.Length);
        SortedDictionary<byte, byte[]>? extended = ExtendedFields(source, fields, before, after, context, resources);
        if (extended == null) return null;
        PublisherBlock[] sizes = fields.Where(field => field.Id == 0x0C).ToArray();
        PublisherBlock[] bolds = fields.Where(field => field.Id == 2).ToArray();
        double nativeSize = after.FontSize * 12700;
        if (size && (after.FontSize > 1000 || nativeSize <= 0 || nativeSize != Math.Round(nativeSize)
                || sizes.Length > 1 || sizes.Any(field => field.Type != 0x22) || fields.Any(field => field.Id == 0x39))
            || bold && (bolds.Length > 1 || bolds.Any(field => field.Type is not (0x02 or 0x0A)) || fields.Any(field => field.Id == 0x37))) return null;
        var missing = new List<byte>();
        if (bold && bolds.Length == 0) missing.Add(2);
        if (size && sizes.Length == 0) missing.Add(0x0C);
        missing.AddRange(extended.Keys.Where(id => fields.All(field => field.Id != id)));
        missing.Sort();
        using var stream = new MemoryStream(span.Length + 8);
        using (var writer = new BinaryWriter(stream, System.Text.Encoding.UTF8, leaveOpen: true)) {
            writer.Write(0U);
            foreach (PublisherBlock field in fields) {
                context.Record();
                foreach (byte id in missing.Where(id => id < field.Id).ToArray()) { Write(id); missing.Remove(id); }
                if (bold && field.Id == 2 || size && field.Id == 0x0C || extended.ContainsKey(field.Id)) Write(field.Id);
                else writer.Write(source.Bytes, field.Offset - 2, field.Length + 2);
            }
            foreach (byte id in missing) Write(id);
            writer.Flush(); stream.Position = 0; writer.Write(checked((uint)stream.Length));
            void Write(byte id) {
                writer.Write(id);
                if (extended.TryGetValue(id, out byte[]? replacement)) writer.Write(replacement, 1, replacement.Length - 1);
                else if (id == 2) writer.Write(after.Bold ? (byte)0x0A : (byte)0x02);
                else { writer.Write((byte)0x22); writer.Write(checked((uint)nativeSize)); }
            }
        }
        return stream.ToArray();
    }
}
