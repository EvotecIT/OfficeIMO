using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherQuillStyleReader {
    private readonly PublisherBinaryData _data;
    private readonly PublisherParseContext _context;
    private readonly PublisherBlockReader _blocks;
    private readonly IReadOnlyList<PublisherQuillChunk> _chunks;
    private readonly PublisherQuillChunk? _text;
    private readonly PublisherQuillStyles _styles;
    private readonly List<string> _fonts = new();
    private readonly List<uint> _colors = new();
    internal PublisherQuillStyleReader(PublisherBinaryData data, PublisherParseContext context,
        IReadOnlyList<PublisherQuillChunk> chunks, IReadOnlyList<OfficeColor?> palette, PublisherQuillChunk? text) {
        _data = data; _context = context; _blocks = new PublisherBlockReader(data, context);
        _chunks = chunks; _text = text; _styles = new PublisherQuillStyles(context, palette);
    }
    internal PublisherQuillStyles Read() {
        PublisherQuillChunk? font = PublisherQuillDirectory.Single(_chunks, "FONT");
        if (font != null) ReadFonts(font);
        PublisherQuillChunk? colors = PublisherQuillDirectory.Single(_chunks, "PL  ");
        if (colors != null) ReadColors(colors);
        _styles.Catalog = ReadCatalog();
        if (_text == null || _text.Length == 0) return _styles;
        foreach (PublisherQuillChunk chunk in _chunks.Where(item => item.Tag is "FDPC" or "FDPP")) ReadRanges(chunk);
        _styles.Characters.Sort((left, right) => left.End.CompareTo(right.End));
        _styles.Paragraphs.Sort((left, right) => left.End.CompareTo(right.End));
        CheckRanges(_styles.Characters.Select(item => item.End));
        CheckRanges(_styles.Paragraphs.Select(item => item.End));
        if (_styles.Characters.Count == 0 || _styles.Paragraphs.Count == 0)
            _context.Add("PUB_TEXT_STYLE_DEFAULTED", "Some native text formatting was unavailable; those runs use the recovered default style.", OfficeConversionLossKind.Approximation, "Quill");
        return _styles;
    }
    private void ReadFonts(PublisherQuillChunk chunk) {
        _data.Range(chunk.Offset, 20, chunk.End);
        int count = _data.Offset(_data.U32(chunk.Offset + 4));
        int position = checked(chunk.Offset + 20 + count * 4);
        _data.Range(chunk.Offset + 20, checked(count * 4), chunk.End);
        for (int i = 0; i < count; i++) {
            _context.Record(); _data.Range(position, 2, chunk.End);
            int bytes = _data.U16(position) * 2;
            _data.Range(position + 2, checked(bytes + 4), chunk.End);
            string name = _data.Utf16(position + 2, bytes).TrimEnd('\0');
            _fonts.Add(string.IsNullOrWhiteSpace(name) ? "Times New Roman" : name);
            position += bytes + 6;
        }
    }
    private void ReadColors(PublisherQuillChunk chunk) {
        _data.Range(chunk.Offset, 12, chunk.End);
        int count = _data.Offset(_data.U32(chunk.Offset)), position = chunk.Offset + 12;
        for (int i = 0; i < count; i++) {
            _context.Record();
            IReadOnlyList<PublisherBlock> fields = _blocks.Chunk(position, chunk.End);
            _colors.Add(PublisherContentsReader.Field(fields, 1)?.Value ?? 0);
            position += _data.Offset(_data.U32(position));
        }
    }
    private void ReadRanges(PublisherQuillChunk chunk) {
        PublisherQuillChunk text = _text!;
        _data.Range(chunk.Offset, 8, chunk.End);
        int count = _data.U16(chunk.Offset);
        _data.Range(chunk.Offset + 8, checked(count * 6), chunk.End);
        for (int i = 0; i < count; i++) {
            _context.Record();
            int absoluteEnd = _data.Offset(_data.U32(chunk.Offset + 8 + i * 4));
            if (absoluteEnd < text.Offset || absoluteEnd > text.End || ((absoluteEnd - text.Offset) & 1) != 0)
                throw new InvalidDataException("Publisher formatting range exceeds its text section.");
            int position = checked(chunk.Offset + _data.U16(chunk.Offset + 8 + count * 4 + i * 2));
            int end = (absoluteEnd - text.Offset) / 2;
            int length = _data.Offset(_data.U32(position));
            var source = new PublisherQuillSourceSpan(position, length);
            if (chunk.Tag == "FDPC") _styles.Characters.Add(new PublisherCharacterRange(end, ReadCharacter(position, chunk.End), source, chunk.Id, i));
            else _styles.Paragraphs.Add(new PublisherParagraphRange(end, ReadParagraph(position, chunk.End), source, chunk.Id, i));
        }
    }
    private void CheckRanges(IEnumerable<int> ends) {
        int last = -1;
        foreach (int end in ends) { if (end <= last) throw new InvalidDataException("Publisher formatting endpoints are not strictly ordered."); last = end; }
    }
    private PublisherCharacterStyle ReadCharacter(int offset, int end) {
        var result = new PublisherCharacterStyle();
        foreach (PublisherBlock field in _blocks.Chunk(offset, end)) {
            switch (field.Id) {
                case 0x02: case 0x37: result.Bold = field.Type != 0x02; break;
                case 0x03: case 0x38: result.Italic = field.Type != 0x02; break;
                case 0x1E: result.Underline = field.Value != 0; break;
                case 0x0C: case 0x39:
                    if (field.Value == 0 || field.Value > 12700 * 1000) throw new InvalidDataException("Invalid Publisher font size.");
                    result.Size = field.Value / 12700D; break;
                case 0x2E: result.Color = field.Value; break;
                case 0x44: result.Color = PublisherContentsReader.Field(_blocks.Children(field), 0)?.Value; break;
                case 0x24:
                    foreach (PublisherBlock font in _blocks.Children(field)) {
                        if (font.Type == 0x88) { result.Font = _blocks.Children(font).FirstOrDefault().Value; break; }
                    }
                    break;
                case 0x0F:
                    result.Baseline = field.Value switch { 1 => OfficeTextBaseline.Superscript, 2 => OfficeTextBaseline.Subscript, _ => OfficeTextBaseline.Normal }; break;
            }
        }
        return result;
    }
}
