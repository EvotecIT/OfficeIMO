using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal static partial class PublisherQuillCharacterStyleWriter {
    // An unspecified replacement retains its selected native source definition.
    internal static bool HasKerningChange(OfficeRichTextRun before, OfficeRichTextRun after) =>
        after.AutomaticPairKerningThreshold.HasValue && before.AutomaticPairKerningThreshold != after.AutomaticPairKerningThreshold;

    internal static bool HasExplicitAttributeChange(OfficeRichTextRun before, OfficeRichTextRun after) =>
        HasLanguageChange(before, after) || HasKerningChange(before, after);

    private static byte[]? KerningField(IReadOnlyList<PublisherBlock> fields, double threshold) {
        double value = Math.Round(threshold * 12700D);
        PublisherBlock[] source = fields.Where(field => field.Id == 0x18).ToArray();
        if (threshold < 0 || threshold > 999.5D || double.IsNaN(value) || threshold != value / 12700D
            || source.Length > 1 || source.Any(field => field.Type != 0x22 || field.Value > 12700 * 999.5D)) return null;
        return PublisherOplWriter.UInt32(0x18, checked((uint)value));
    }

    // Retained CRs have no visible run to trigger the ordinary character codec.
    // Apply only the requested nullable attributes, leaving their other fields intact.
    internal static byte[]? RewriteExplicitAttributes(byte[] style, OfficeRichTextRun before, OfficeRichTextRun after,
        PublisherParseContext context) {
        if (!HasExplicitAttributeChange(before, after)) return style;
        var source = new PublisherBinaryData(style, "Quill/FDPC");
        var fields = new PublisherBlockReader(source, context).Chunk(0, style.Length);
        var changes = new SortedDictionary<byte, byte[]>();
        if (HasLanguageChange(before, after)) {
            if (!LanguageFieldIsQualified(fields) || after.LanguageId!.Value > ushort.MaxValue) return null;
            changes.Add(0x12, PublisherOplWriter.UInt32(0x12, after.LanguageId.Value));
        }
        if (HasKerningChange(before, after)) {
            byte[]? field = KerningField(fields, after.AutomaticPairKerningThreshold!.Value);
            if (field == null) return null;
            changes.Add(0x18, field);
        }
        return PublisherOplWriter.Rewrite(source, 0, style.Length, fields, changes, context);
    }
}
