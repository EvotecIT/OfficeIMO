namespace OfficeIMO.DjVu;

internal sealed class DjVuContainerReader {
    private readonly DjVuReadBudget _budget;
    private long _sourceBytes;
    internal DjVuChunk Root { get; private set; } = null!;
    internal DjVuContainerReader(DjVuReadBudget budget) => _budget = budget;

    internal List<DjVuComponent> Read(byte[] source) {
        _sourceBytes = source.Length;
        DjVuChunk root = ParseFile(source, true);
        Root = root;
        if (root.FormType == "DJVU" || root.FormType == "PM44" || root.FormType == "BM44")
            return new List<DjVuComponent> { new DjVuComponent("page-1", "page-1", "1", 1, root) };
        if (root.FormType != "DJVM") throw new NotSupportedException("The input is not a supported DjVu document.");
        if (root.Children.Count == 0 || root.Children[0].Id != "DIRM") throw new InvalidDataException("DjVu book has no leading DIRM directory.");
        return ReadDirectory(root, root.Children[0]);
    }

    private DjVuChunk ParseFile(byte[] source, bool requireSignature) {
        bool signature = source.Length >= 4 && source[0] == 65 && source[1] == 84 && source[2] == 38 && source[3] == 84;
        int start = signature ? 4 : 0;
        DjVuChunk root = ParseChunk(source, start, source.Length, 1, out int next);
        if (root.Id != "FORM" || next != source.Length) throw new InvalidDataException("Invalid DjVu root container or trailing bytes.");
        if (requireSignature && !signature && root.FormType != "PM44" && root.FormType != "BM44")
            throw new InvalidDataException("Missing DjVu AT&T signature.");
        return root;
    }

    private DjVuChunk ParseChunk(byte[] source, int start, int end, int depth, out int next) {
        _budget.Chunk(depth);
        if (end - start < 8) throw new InvalidDataException("Truncated IFF chunk header.");
        string id = DjVuBinary.Id(source, start);
        uint encodedSize = DjVuBinary.U32(source, start + 4);
        if (encodedSize > end - start - 8) throw new InvalidDataException("IFF chunk exceeds its containing bytes.");
        int length = (int)encodedSize, offset = start + 8;
        int dataEnd = offset + length;
        next = dataEnd;
        if ((length & 1) != 0 && next < end) next++;
        var children = new List<DjVuChunk>();
        string? formType = null;
        if (id == "FORM") {
            if (length < 4) throw new InvalidDataException("IFF FORM has no subtype.");
            formType = DjVuBinary.Id(source, offset);
            int child = offset + 4;
            while (child < dataEnd) {
                children.Add(ParseChunk(source, child, dataEnd, depth + 1, out int childEnd));
                child = childEnd;
            }
        }
        return new DjVuChunk(source, id, formType, offset, length, children);
    }

    private List<DjVuComponent> ReadDirectory(DjVuChunk root, DjVuChunk directory) {
        byte[] source = directory.Source;
        int start = directory.Offset, end = start + directory.Length;
        if (directory.Length < 3) throw new InvalidDataException("Truncated DjVu directory.");
        int flags = source[start++];
        if ((flags & 127) != 1) throw new NotSupportedException("Unsupported DjVu directory version.");
        bool bundled = (flags & 128) != 0;
        int count = DjVuBinary.U16(source, start); start += 2;
        if (count == 0) throw new InvalidDataException("Empty DjVu directory.");
        if (count > _budget.Options.MaxComponents) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxComponents));
        var offsets = new uint[count];
        if (bundled) {
            if (count > (end - start) / 4) throw new InvalidDataException("Truncated DjVu component offsets.");
            for (int i = 0; i < count; i++, start += 4) offsets[i] = DjVuBinary.U32(source, start);
        }
        byte[] data = BzzDecoder.Decode(source, start, end - start, _budget);
        if (count > data.Length / 4) throw new InvalidDataException("Truncated DjVu component records.");
        int position = count * 4, pageCount = 0;
        var components = new List<DjVuComponent>(count);
        var ids = new HashSet<string>(StringComparer.Ordinal);
        var usedOffsets = new HashSet<uint>();
        var forms = root.Children.Where(c => c.Id == "FORM").ToDictionary(c => (uint)c.HeaderOffset);
        for (int i = 0; i < count; i++) {
            _budget.Cancellation.ThrowIfCancellationRequested();
            int size = DjVuBinary.U24(data, i * 3);
            int recordFlags = data[count * 3 + i];
            int kind = recordFlags & 63;
            if (kind > 2) throw new NotSupportedException("Unsupported DjVu component kind.");
            string id = ReadString(data, ref position);
            if (id.Length == 0 || !ids.Add(id)) throw new InvalidDataException("Empty or duplicate DjVu component identity.");
            string name = (recordFlags & 128) != 0 ? ReadString(data, ref position) : id;
            string title = (recordFlags & 64) != 0 ? ReadString(data, ref position) : id;
            if (kind == 1 && ++pageCount > _budget.Options.MaxPages) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxPages));
            DjVuChunk form;
            if (bundled) {
                if (!usedOffsets.Add(offsets[i]) || !forms.TryGetValue(offsets[i], out form!)) throw new InvalidDataException("DjVu directory points outside a unique component FORM.");
                if (size != form.Length + 8) throw new InvalidDataException("DjVu component size does not match its FORM.");
            } else {
                var resolver = _budget.Options.ComponentResolver ?? throw new NotSupportedException("Indirect DjVu documents require an explicit component resolver.");
                byte[] resolved = resolver(id, _budget.Cancellation) ?? throw new InvalidDataException("DjVu component resolver returned no bytes.");
                if (resolved.LongLength > _budget.Options.MaxSourceBytes - _sourceBytes) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxSourceBytes));
                _sourceBytes += resolved.Length;
                var owned = (byte[])resolved.Clone();
                form = ParseFile(owned, false);
            }
            string expectedType = kind == 0 ? "DJVI" : kind == 1 ? "DJVU" : "THUM";
            if (form.FormType != expectedType) throw new InvalidDataException("DjVu component kind does not match its FORM subtype.");
            components.Add(new DjVuComponent(id, name, title, kind, form));
        }
        if (position != data.Length || pageCount == 0) throw new InvalidDataException("Invalid DjVu directory records.");
        if (bundled && usedOffsets.Count != forms.Count) throw new InvalidDataException("DjVu book contains unlisted component forms.");
        return components;
    }

    private static string ReadString(byte[] data, ref int offset) {
        int start = offset;
        while (offset < data.Length && data[offset] != 0) offset++;
        if (offset == data.Length) throw new InvalidDataException("Unterminated DjVu component identity.");
        string value = DjVuBinary.Utf8.GetString(data, start, offset - start);
        offset++;
        return value;
    }
}
