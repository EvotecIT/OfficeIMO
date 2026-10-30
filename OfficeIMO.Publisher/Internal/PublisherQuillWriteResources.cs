using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Internal;

/// <summary>Appends explicit font/color entries without changing existing native identities or opaque entry metadata.</summary>
internal sealed class PublisherQuillWriteResources {
    private readonly PublisherQuillSourceLayout _layout;
    private readonly PublisherParseContext _context;
    private readonly List<byte[]> _fonts = new(), _colors = new();
    private readonly List<string> _fontNames = new();
    private readonly List<uint?> _plainColors = new();
    private byte[]? _fontHeader, _colorHeader;
    private bool _fontsChanged, _colorsChanged;

    internal PublisherQuillWriteResources(PublisherQuillSourceLayout layout, PublisherParseContext context) {
        _layout = layout; _context = context;
    }

    internal bool Font(string name, out uint index) {
        index = 0;
        if (name.Length is < 1 or > 32 || string.IsNullOrWhiteSpace(name) || name.Any(char.IsControl)) return false;
        byte[] encoded;
        try { encoded = new UnicodeEncoding(false, false, true).GetBytes(name); }
        catch (EncoderFallbackException) { return false; }
        if (_fontHeader == null && !ReadFonts()) return false;
        int found = _fontNames.IndexOf(name);
        if (found >= 0) { index = (uint)found; return true; }
        if (_fonts.Count >= ushort.MaxValue) return false;
        _context.Record(); index = (uint)_fonts.Count;
        var entry = new byte[checked(encoded.Length + 6)];
        entry[0] = (byte)name.Length; entry[1] = (byte)(name.Length >> 8);
        Buffer.BlockCopy(encoded, 0, entry, 2, encoded.Length);
        // Publisher resolves the name and replaces the saved low-word ID on
        // load. The two additional hint bytes are unspecified for a new family.
        _fonts.Add(entry); _fontNames.Add(name); _fontsChanged = true; return true;
    }

    internal bool Color(OfficeColor color, out uint index) {
        index = 0;
        if (color.A != 255 || _colorHeader == null && !ReadColors()) return false;
        uint rgb = (uint)(color.R | color.G << 8 | color.B << 16);
        int found = _plainColors.FindIndex(value => value == rgb);
        if (found >= 0) { index = (uint)found; return true; }
        _context.Record(); index = (uint)_colors.Count;
        // Native plain RGB entries: color field1 and the empty field6 envelope.
        byte[] entry = { 18, 0, 0, 0, 1, 0x22, 0, 0, 0, 0, 6, 0x82, 6, 0, 0, 0, 0, 0 };
        Put(entry, 6, rgb); _colors.Add(entry); _plainColors.Add(rgb); _colorsChanged = true; return true;
    }

    internal IReadOnlyDictionary<string, byte[]> Replacements() {
        var result = new Dictionary<string, byte[]>();
        if (_fontsChanged) {
            int size = checked(20 + _fonts.Count * 4 + _fonts.Sum(entry => entry.Length));
            var output = new byte[size]; Buffer.BlockCopy(_fontHeader!, 0, output, 0, 20);
            Put(output, 0, (uint)(size - 20)); Put(output, 4, (uint)_fonts.Count);
            int position = checked(20 + _fonts.Count * 4);
            for (int i = 0; i < _fonts.Count; i++) {
                _context.Record(); Put(output, 20 + i * 4, (uint)(position - 20));
                Buffer.BlockCopy(_fonts[i], 0, output, position, _fonts[i].Length); position += _fonts[i].Length;
            }
            result.Add("FONT", output);
        }
        if (_colorsChanged) {
            var output = new byte[checked(12 + _colors.Sum(entry => entry.Length))];
            Buffer.BlockCopy(_colorHeader!, 0, output, 0, 12); Put(output, 0, (uint)_colors.Count);
            int position = 12;
            foreach (byte[] entry in _colors) {
                _context.Record(); Buffer.BlockCopy(entry, 0, output, position, entry.Length); position += entry.Length;
            }
            result.Add("PL  ", output);
        }
        return result;
    }

    private bool ReadFonts() {
        PublisherQuillChunk chunk = _layout.Find("FONT"); PublisherBinaryData data = _layout.Data;
        if (chunk.Length < 20 || data.U32(chunk.Offset) != chunk.Length - 20 || data.U32(chunk.Offset + 8) != 4
            || data.U32(chunk.Offset + 12) != 4 || data.U32(chunk.Offset + 16) != 0) return false;
        uint count = data.U32(chunk.Offset + 4);
        if (count > ushort.MaxValue || 20L + count * 4 > chunk.Length) return false;
        int position = checked(chunk.Offset + 20 + (int)count * 4);
        for (int i = 0; i < count; i++) {
            _context.Record();
            if (position > chunk.End - 2 || data.U32(chunk.Offset + 20 + i * 4) != position - chunk.Offset - 20) return false;
            int bytes = data.U16(position) * 2, length = checked(bytes + 6);
            if (length > chunk.End - position) return false;
            _fontNames.Add(data.Utf16(position + 2, bytes).TrimEnd('\0'));
            _fonts.Add(Copy(position, length)); position += length;
        }
        if (position != chunk.End) return false;
        _fontHeader = Copy(chunk.Offset, 20); return true;
    }

    private bool ReadColors() {
        PublisherQuillChunk chunk = _layout.Find("PL  "); PublisherBinaryData data = _layout.Data;
        if (chunk.Length < 12 || data.U32(chunk.Offset + 4) != 40 || data.U32(chunk.Offset + 8) != 0) return false;
        uint count = data.U32(chunk.Offset); int position = chunk.Offset + 12;
        if (count > (chunk.Length - 12) / 4) return false;
        var blocks = new PublisherBlockReader(data, _context);
        for (int i = 0; i < count; i++) {
            _context.Record();
            if (position > chunk.End - 4) return false;
            int length = data.Offset(data.U32(position));
            if (length < 4 || length > chunk.End - position) return false;
            IReadOnlyList<PublisherBlock> fields = blocks.Chunk(position, chunk.End);
            PublisherBlock[] value = fields.Where(field => field.Id == 1).ToArray();
            if (value.Length != 1 || value[0].Type != 0x22) return false;
            byte[] raw = Copy(position, length);
            bool plain = length == 18 && raw[4] == 1 && raw[5] == 0x22
                && raw.Skip(10).SequenceEqual(new byte[] { 6, 0x82, 6, 0, 0, 0, 0, 0 }) && (value[0].Value & 0xFF000000) == 0;
            _colors.Add(raw); _plainColors.Add(plain ? value[0].Value : null); position += length;
        }
        if (position != chunk.End) return false;
        _colorHeader = Copy(chunk.Offset, 12); return true;
    }

    private byte[] Copy(int offset, int length) {
        var result = new byte[length]; Buffer.BlockCopy(_layout.Data.Bytes, offset, result, 0, length); return result;
    }
    private static void Put(byte[] bytes, int offset, uint value) {
        for (int i = 0; i < 4; i++) bytes[offset + i] = (byte)(value >> (i * 8));
    }
}
