using OfficeIMO.Drawing;
using static OfficeIMO.Publisher.Internal.PublisherQuillOplWriter;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Changes qualified paragraph fields without rewriting their style selector or unrelated native metadata.</summary>
internal static class PublisherQuillParagraphStyleWriter {
    internal static bool HasChanges(OfficeRichTextParagraph before, OfficeRichTextParagraph after) => before.Alignment != after.Alignment
        || before.LineHeight != after.LineHeight || before.LineHeightFactor != after.LineHeightFactor
        || !before.Margins.Equals(after.Margins) || before.Indent != after.Indent || !TabsEqual(before.TabStops, after.TabStops);

    internal static byte[]? Rewrite(PublisherBinaryData source, PublisherQuillSourceSpan span,
        OfficeRichTextParagraph before, OfficeRichTextParagraph after, PublisherParseContext context) {
        if (!HasChanges(before, after)) {
            var copy = new byte[span.Length]; Buffer.BlockCopy(source.Bytes, span.Offset, copy, 0, span.Length); return copy;
        }
        OfficeTextTabStops? tabs = after.TabStops;
        bool aligned = after.Alignment is OfficeTextAlignment.Center or OfficeTextAlignment.Right;
        if (tabs == null || tabs.DefaultInterval != 36 || tabs.Origin != -after.Margins.Left || tabs.AlignWithParagraph != aligned
            || tabs.Stops.Any(stop => stop.Alignment != OfficeTextTabAlignment.Left || stop.Character != "."
                || stop.LeaderText != null || stop.LineLeader != null || stop.LeaderStyle != null)) return null;
        var blocks = new PublisherBlockReader(source, context);
        IReadOnlyList<PublisherBlock> fields = blocks.Chunk(span.Offset, span.Offset + span.Length);
        var changes = new SortedDictionary<byte, byte[]>();
        if (before.Alignment != after.Alignment) {
            if (fields.Any(field => field.Id == 4 && field.Value is not (0 or 1 or 2 or 6))) return null;
            ushort alignment = after.Alignment switch {
                OfficeTextAlignment.Left => 0, OfficeTextAlignment.Right => 1, OfficeTextAlignment.Center => 2,
                OfficeTextAlignment.Justify => 6, _ => ushort.MaxValue
            };
            if (alignment == ushort.MaxValue || !Set(4, 0x12, UInt16(4, alignment))) return null;
        }
        if (before.Margins.Left != after.Margins.Left || before.Indent != after.Indent) {
            if (after.Indent.FirstLineOffset != 0 && after.Indent.ContinuationLineOffset != 0) return null;
            double left = after.Margins.Left + after.Indent.ContinuationLineOffset;
            double first = after.Indent.FirstLineOffset - after.Indent.ContinuationLineOffset;
            if (left != before.Margins.Left + before.Indent.ContinuationLineOffset && !Point(0x0D, left, true)
                || first != before.Indent.FirstLineOffset - before.Indent.ContinuationLineOffset && !Point(0x0C, first, true)) return null;
        }
        if (before.Margins.Right != after.Margins.Right && !Point(0x0E, after.Margins.Right, true)
            || before.Margins.Top != after.Margins.Top && !Point(0x12, after.Margins.Top, false)
            || before.Margins.Bottom != after.Margins.Bottom && !Point(0x13, after.Margins.Bottom, false)) return null;
        if (before.LineHeight != after.LineHeight || before.LineHeightFactor != after.LineHeightFactor) {
            if (fields.Any(field => field.Id == 0x34 && (field.Value & 7) is not (1 or 2))) return null;
            double? amount = after.LineHeight ?? after.LineHeightFactor;
            double unit = after.LineHeight.HasValue ? 8D * 12700 : 96D * 12700;
            uint flag = after.LineHeight.HasValue ? 1U : 2U;
            if (!amount.HasValue || !Scale(amount.Value, unit, 1, uint.MaxValue - flag, out long value) || (value & 7) != 0
                || !Set(0x34, 0x22, UInt32(0x34, (uint)value + flag))) return null;
        }
        if (!TabContentEqual(before.TabStops, tabs)) {
            if (!KnownTabs() || !WriteTabs(out byte[]? replacement) || !Set(0x32, 0x82, replacement!)) return null;
        }
        return PublisherQuillOplWriter.Rewrite(source, span, fields, changes, context);

        bool Set(byte id, byte type, byte[] replacement) {
            PublisherBlock[] old = fields.Where(field => field.Id == id).ToArray();
            if (old.Length > 1 || old.Any(field => field.Type != type)) return false;
            changes.Add(id, replacement); return true;
        }
        bool Point(byte id, double value, bool signed) {
            if (!Scale(value, 12700, signed ? int.MinValue : 0, signed ? int.MaxValue : uint.MaxValue, out long scaled)) return false;
            return Set(id, 0x22, UInt32(id, unchecked((uint)scaled)));
        }
        bool KnownTabs() {
            PublisherBlock[] old = fields.Where(field => field.Id == 0x32).ToArray();
            if (old.Length == 0) return true;
            if (old.Length != 1 || old[0].Type != 0x82) return false;
            IReadOnlyList<PublisherBlock> children = blocks.Children(old[0]);
            if (children.Count != 2 || children[0].Id != 0x27 || children[0].Type != 0x1A
                || children[1].Id != 0x28 || children[1].Type != 0x8A) return false;
            IReadOnlyList<PublisherBlock> entries = blocks.Children(children[1]);
            if (children[0].Value != entries.Count || entries.Count > 256) return false;
            for (int i = 0; i < entries.Count; i++) {
                PublisherBlock entry = entries[i];
                if (entry.Id != 0 && entry.Id != i || entry.Type != 0x88) return false;
                IReadOnlyList<PublisherBlock> value = blocks.Children(entry);
                if (value.Count != 1 || value[0].Id != 0 || value[0].Type is not (0x20 or 0x22) || value[0].Value > int.MaxValue) return false;
            }
            return true;
        }
        bool WriteTabs(out byte[]? replacement) {
            var entries = new List<byte[]>(); replacement = null;
            foreach (OfficeTextTabStop stop in tabs.Stops) {
                context.Record();
                if (!Scale(stop.Position, 12700, 0, int.MaxValue, out long position)) return false;
                entries.Add(Nested((byte)entries.Count, 0x88, UInt32(0, (uint)position, 0x20)));
            }
            replacement = Nested(0x32, 0x82, UInt16(0x27, (ushort)entries.Count, 0x1A), Nested(0x28, 0x8A, entries.ToArray())); return true;
        }
    }

    // Compare decoded units rather than binary floating-point products: values
    // such as 1.21 points can multiply to just below an integral EMU.
    private static bool Scale(double value, double unit, long minimum, long maximum, out long result) {
        double scaled = Math.Round(value * unit); result = 0;
        if (scaled < minimum || scaled > maximum || double.IsNaN(scaled) || scaled / unit != value) return false;
        result = (long)scaled; return true;
    }

    internal static bool TabsEqual(OfficeTextTabStops? a, OfficeTextTabStops? b) => ReferenceEquals(a, b)
        || a != null && b != null && a.Origin == b.Origin && a.AlignWithParagraph == b.AlignWithParagraph && TabContentEqual(a, b);
    private static bool TabContentEqual(OfficeTextTabStops? a, OfficeTextTabStops? b) => ReferenceEquals(a, b)
        || a != null && b != null && a.DefaultInterval == b.DefaultInterval
        && a.Stops.Count == b.Stops.Count && a.Stops.Zip(b.Stops, (x, y) => x.Position == y.Position && x.Alignment == y.Alignment
            && x.Character == y.Character && x.LeaderText == y.LeaderText && Equals(x.LineLeader, y.LineLeader) && Equals(x.LeaderStyle, y.LeaderStyle)).All(value => value);
}
