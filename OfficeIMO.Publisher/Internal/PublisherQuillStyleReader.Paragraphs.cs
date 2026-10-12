using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherQuillStyleReader {
    private PublisherParagraphStyle ReadParagraph(int offset, int end) {
        var result = new PublisherParagraphStyle();
        foreach (PublisherBlock field in _blocks.Chunk(offset, end)) {
            switch (field.Id) {
                case 0x19: result.DefaultStyleIndex = field.Value; break;
                case 0x04:
                    result.Alignment = (field.Value & 255) switch { 2 => OfficeTextAlignment.Center, 1 => OfficeTextAlignment.Right,
                        6 => OfficeTextAlignment.Justify, _ => OfficeTextAlignment.Left }; break;
                case 0x34:
                    if ((field.Value & 1) != 0 && field.Value > 1) result.LineHeight = (field.Value - 1) / (8D * 12700);
                    else if ((field.Value & 2) != 0 && field.Value > 2) result.LineHeightFactor = (field.Value - 2) / (96D * 12700);
                    break;
                case 0x12: result.Before = field.Value / 12700D; break;
                case 0x13: result.After = field.Value / 12700D; break;
                case 0x0D: result.Left = unchecked((int)field.Value) / 12700D; break;
                case 0x0E: result.Right = unchecked((int)field.Value) / 12700D; break;
                case 0x0C: result.FirstLine = unchecked((int)field.Value) / 12700D; break;
                case 0x02:
                    if (field.Value > 12700 * 1000) throw new InvalidDataException("Invalid Publisher list-label size.");
                    result.LabelSize = field.Value / 12700D; break;
                case 0x03: result.LabelFont = field.Value; break;
                case 0x06: result.LabelTextPosition = field.Value / 12700D; break;
                case 0x32: result.Tabs = ReadTabs(field); break;
                case 0x57:
                    IReadOnlyList<PublisherBlock> list = _blocks.Children(field);
                    uint type = PublisherContentsReader.Field(list, 0)?.Value ?? 0xFF;
                    uint bullet = PublisherContentsReader.Field(list, 1)?.Value ?? 0;
                    result.List = new PublisherListStyle(type, bullet);
                    if (bullet == 0 && type != 0xFF) _context.Add("PUB_NUMBERED_LIST_UNASSESSED",
                        "A native numbered-list definition is retained as paragraph text without an inferred numbering sequence.", OfficeConversionLossKind.Unassessed, "Quill/FDPP");
                    break;
                case 0x08: case 0x2C: case 0x2D:
                    if (field.Value != 0) _context.Add("PUB_DROP_CAP_UNASSESSED", "Native drop-cap geometry requires additional text projection.",
                        OfficeConversionLossKind.Unassessed, "Quill/FDPP"); break;
            }
        }
        return result;
    }

    private IReadOnlyList<OfficeTextTabStop> ReadTabs(PublisherBlock field) {
        IReadOnlyList<PublisherBlock> fields = _blocks.Children(field);
        uint? count = PublisherContentsReader.Field(fields, 0x27)?.Value;
        PublisherBlock? array = PublisherContentsReader.Field(fields, 0x28);
        var result = new List<OfficeTextTabStop>();
        var positions = new HashSet<uint>();
        if (array.HasValue) {
            foreach (PublisherBlock entry in _blocks.Children(array.Value)) {
                _context.Record();
                if (entry.Id != 0 || entry.Type != 0x88) throw new InvalidDataException("Invalid Publisher tab-stop entry.");
                IReadOnlyList<PublisherBlock> properties = _blocks.Children(entry);
                PublisherBlock amount = PublisherContentsReader.Field(properties, 0)
                    ?? throw new InvalidDataException("Publisher tab stop has no position.");
                if (amount.Type is not (0x20 or 0x22) || unchecked((int)amount.Value) < 0 || !positions.Add(amount.Value)
                    || result.Count >= 256) throw new InvalidDataException("Publisher tab positions are invalid, duplicate, or exceed 256 stops.");
                result.Add(new OfficeTextTabStop(amount.Value / 12700D));
                if (properties.Any(property => property.Id != 0)) _context.Add("PUB_TAB_DETAIL_UNASSESSED",
                    "Additional native tab alignment or leader fields are not projected by this tab profile.", OfficeConversionLossKind.Unassessed, "Quill/FDPP");
            }
        }
        if (count.HasValue && count.Value != result.Count) throw new InvalidDataException("Publisher tab-stop count disagrees with its native array.");
        _context.Add("PUB_TAB_LAYOUT_APPROXIMATED", "Native tab positions use the shared measured left-tab grid. Undeclared stops use the shared 36-point interval.",
            OfficeConversionLossKind.Approximation, "Quill/FDPP");
        return result;
    }
}
