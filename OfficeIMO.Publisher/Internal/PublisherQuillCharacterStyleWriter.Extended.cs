using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillCharacterStyleWriter {
    internal static bool HasExtendedChanges(OfficeRichTextRun before, OfficeRichTextRun after) => before.Color != after.Color
        || before.FontFamily != after.FontFamily || before.Italic != after.Italic || before.UnderlineStyle != after.UnderlineStyle
        || before.StrikethroughStyle != after.StrikethroughStyle || before.Baseline != after.Baseline;

    private static SortedDictionary<byte, byte[]>? ExtendedFields(PublisherBinaryData source, IReadOnlyList<PublisherBlock> fields,
        OfficeRichTextRun before, OfficeRichTextRun after, PublisherParseContext context, PublisherQuillWriteResources? resources) {
        var blocks = new PublisherBlockReader(source, context);
        var changes = new SortedDictionary<byte, byte[]>();
        if (before.Italic != after.Italic) {
            if (!Flag(3) || fields.Any(field => field.Id == 0x38)) return null;
            changes.Add(3, Boolean(3, after.Italic));
        }
        if (before.StrikethroughStyle != after.StrikethroughStyle) {
            if (!Flag(0x10) || after.StrikethroughStyle is not (OfficeTextDecorationStyle.None or OfficeTextDecorationStyle.Single)) return null;
            changes.Add(0x10, Boolean(0x10, after.Strikethrough));
        }
        if (before.UnderlineStyle != after.UnderlineStyle) {
            if (!Scalar(0x1E, 0x12)) return null;
            ushort value = after.UnderlineStyle switch {
                OfficeTextDecorationStyle.None => 0, OfficeTextDecorationStyle.Single => 1, OfficeTextDecorationStyle.Words => 2,
                OfficeTextDecorationStyle.Double => 3, OfficeTextDecorationStyle.Dotted => 4,
                OfficeTextDecorationStyle.Dashed => 7, OfficeTextDecorationStyle.Wavy => 11, _ => ushort.MaxValue
            };
            if (value == ushort.MaxValue) return null;
            changes.Add(0x1E, UInt16(0x1E, value));
        }
        if (before.Baseline != after.Baseline) {
            if (!Scalar(0x0F, 0x12)) return null;
            changes.Add(0x0F, UInt16(0x0F, (ushort)after.Baseline));
        }
        if (before.FontFamily != after.FontFamily) {
            if (resources == null || !FontField() || !resources.Font(after.FontFamily, out uint index)) return null;
            changes.Add(0x24, Nested(0x24, 0x8A, ScriptIds.Select(id => Nested(id, 0x88, FontIndex((ushort)index))).ToArray()));
        }
        if (before.Color != after.Color) {
            if (resources == null || !ColorFields() || !resources.Color(after.Color, out uint index)) return null;
            changes.Add(0x44, Nested(0x44, 0x8A, UInt32(0, index)));
            changes.Add(0x58, Nested(0x58, 0x8A, UInt16(0, 1), UInt32(1, index), UInt32(2, 100000)));
        }
        return changes;

        bool Scalar(byte id, byte type) {
            PublisherBlock[] selected = fields.Where(field => field.Id == id).ToArray();
            return selected.Length <= 1 && selected.All(field => field.Type == type);
        }
        bool Flag(byte id) {
            PublisherBlock[] selected = fields.Where(field => field.Id == id).ToArray();
            return selected.Length <= 1 && selected.All(field => field.Type is 0x02 or 0x0A);
        }
        bool FontField() {
            PublisherBlock[] selected = fields.Where(field => field.Id == 0x24).ToArray();
            if (selected.Length == 0) return true;
            if (selected.Length != 1 || selected[0].Type != 0x8A) return false;
            IReadOnlyList<PublisherBlock> scripts = blocks.Children(selected[0]);
            if (!scripts.Select(field => field.Id).SequenceEqual(ScriptIds) || scripts.Any(field => field.Type != 0x88)) return false;
            uint? index = null;
            foreach (PublisherBlock script in scripts) {
                IReadOnlyList<PublisherBlock> value = blocks.Children(script);
                if (value.Count != 1 || value[0].Id != 0 || value[0].Type != 0x18 || index.HasValue && index != value[0].Value) return false;
                index = value[0].Value;
            }
            return true;
        }
        bool ColorFields() {
            if (fields.Any(field => field.Id == 0x2E)) return false;
            PublisherBlock[] ink = fields.Where(field => field.Id == 0x44).ToArray(), fill = fields.Where(field => field.Id == 0x58).ToArray();
            if (ink.Length == 0 && fill.Length == 0) return true;
            if (ink.Length != 1 || fill.Length != 1 || ink[0].Type != 0x8A || fill[0].Type != 0x8A) return false;
            IReadOnlyList<PublisherBlock> a = blocks.Children(ink[0]), b = blocks.Children(fill[0]);
            return a.Count == 1 && a[0].Id == 0 && a[0].Type == 0x22 && b.Count == 3
                && b[0].Id == 0 && b[0].Type == 0x12 && b[0].Value == 1
                && b[1].Id == 1 && b[1].Type == 0x22 && b[1].Value == a[0].Value
                && b[2].Id == 2 && b[2].Type == 0x22 && b[2].Value == 100000;
        }
    }

    private static readonly byte[] ScriptIds = { 0, 1, 2, 3, 4, 5, 0x2A, 0x2C };
    private static byte[] Boolean(byte id, bool value) => new[] { id, value ? (byte)0x0A : (byte)0x02 };
    private static byte[] UInt16(byte id, ushort value) => new[] { id, (byte)0x12, (byte)value, (byte)(value >> 8) };
    private static byte[] FontIndex(ushort value) => new[] { (byte)0, (byte)0x18, (byte)value, (byte)(value >> 8) };
    private static byte[] UInt32(byte id, uint value) => new[] { id, (byte)0x22, (byte)value, (byte)(value >> 8), (byte)(value >> 16), (byte)(value >> 24) };
    private static byte[] Nested(byte id, byte type, params byte[][] children) {
        int length = checked(4 + children.Sum(child => child.Length)); var bytes = new byte[length + 2]; bytes[0] = id; bytes[1] = type;
        for (int i = 0; i < 4; i++) bytes[2 + i] = (byte)(length >> (i * 8));
        int position = 6;
        foreach (byte[] child in children) { Buffer.BlockCopy(child, 0, bytes, position, child.Length); position += child.Length; }
        return bytes;
    }
}
