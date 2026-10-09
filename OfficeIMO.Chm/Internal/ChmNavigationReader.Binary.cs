namespace OfficeIMO.Chm;

internal sealed partial class ChmNavigationReader {
    private IReadOnlyList<ChmNavigationItem> ReadBinaryContents(byte[] data) {
        ChmBinary.Range(data, 0, 16);
        int first = ChmBinary.Index(ChmBinary.U32(data, 0));
        uint entryCount = ChmBinary.U32(data, 8);
        if (entryCount == 0) return Array.Empty<ChmNavigationItem>();
        if (first < 16) throw ChmBinary.Error("CONTENTS", "The compiled contents root overlaps its header.");
        var roots = new List<Item>();
        var pending = new Stack<(int Offset, List<Item> Output, int Depth)>();
        var visited = new HashSet<int>();
        pending.Push((first, roots, 1));
        byte[]? names = _book.FindEntry("/#STRINGS")?.GetBytes();
        while (pending.Count != 0) {
            var current = pending.Pop();
            _token.ThrowIfCancellationRequested();
            if (!visited.Add(current.Offset)) throw ChmBinary.Error("CONTENTS", "The compiled contents contain a cycle or shared node.");
            ChmBinary.Range(data, current.Offset, 20);
            Item item = NewItem(current.Depth);
            uint properties = ChmBinary.U32(data, current.Offset + 4), value = ChmBinary.U32(data, current.Offset + 8);
            if ((properties & 8) != 0) {
                var topic = Topic(value); item.Name = topic.Title; item.Links.Add(Link(topic.Target, "/", topic.Title));
            } else {
                if (names == null) throw ChmBinary.Error("CONTENTS", "Compiled contents headings require #STRINGS.");
                item.Name = ChmBinary.CString(names, ChmBinary.Index(value), _encoding, _options.MaxPathLength);
            }
            current.Output.Add(item);
            int sibling = ChmBinary.Index(ChmBinary.U32(data, current.Offset + 16));
            if (sibling != 0) pending.Push((sibling, current.Output, current.Depth));
            if ((properties & 4) != 0) {
                ChmBinary.Range(data, current.Offset, 28);
                int child = ChmBinary.Index(ChmBinary.U32(data, current.Offset + 20));
                if (child != 0) pending.Push((child, item.Children, current.Depth + 1));
            }
        }
        return roots.Select(item => item.Freeze()).ToList().AsReadOnly();
    }

    private IReadOnlyList<ChmNavigationItem> ReadBinaryIndex(byte[] data) {
        const int headerLength = 76;
        ChmBinary.Range(data, 0, headerLength);
        if (data[0] != 0x3B || data[1] != 0x29) throw ChmBinary.Error("INDEX", "The compiled keyword index has an invalid signature.");
        int blockLength = ChmBinary.U16(data, 4);
        int lastListing = ChmBinary.Index(ChmBinary.U32(data, 26));
        int blockCount = ChmBinary.Index(ChmBinary.U32(data, 38));
        if (blockLength < 12 || blockCount <= lastListing || blockCount > (data.Length - headerLength) / blockLength)
            throw ChmBinary.Error("INDEX", "The compiled keyword index dimensions are invalid.");
        var roots = new List<Item>();
        var parents = new List<Item>();
        for (int block = 0; block <= lastListing; block++) {
            _token.ThrowIfCancellationRequested();
            int start = headerLength + block * blockLength;
            int free = ChmBinary.U16(data, start), count = ChmBinary.U16(data, start + 2);
            if (free > blockLength - 12) throw ChmBinary.Error("INDEX", "The keyword index listing free-space size is invalid.");
            int end = start + blockLength - free;
            int position = start + 12;
            for (int record = 0; record < count; record++) {
                string fullName = ReadWideString(data, ref position, end);
                Require(16);
                int seeAlso = ChmBinary.U16(data, position), depth = ChmBinary.U16(data, position + 2);
                uint charIndex = ChmBinary.U32(data, position + 4), pairs = ChmBinary.U32(data, position + 12);
                position += 16;
                if (depth > parents.Count || charIndex > fullName.Length || pairs > _options.MaxNavigationItems)
                    throw ChmBinary.Error("INDEX", "The keyword depth, name offset, or target count is invalid.");
                Item item = NewItem(depth + 1);
                item.Name = fullName.Substring((int)charIndex).TrimStart(' ', ',');
                if (seeAlso == 2) item.SeeAlso.Add(ReadWideString(data, ref position, end));
                else if (seeAlso == 0) {
                    for (uint pair = 0; pair < pairs; pair++) {
                        _token.ThrowIfCancellationRequested(); Require(4);
                        var topic = Topic(ChmBinary.U32(data, position)); position += 4;
                        item.Links.Add(Link(topic.Target, "/", topic.Title));
                    }
                } else throw ChmBinary.Error("INDEX", "An index entry uses an unsupported target type.");
                Require(8); position += 8;
                if (parents.Count > depth) parents.RemoveRange(depth, parents.Count - depth);
                if (depth == 0) roots.Add(item); else parents[depth - 1].Children.Add(item);
                parents.Add(item);
            }
            if (position != end) throw ChmBinary.Error("INDEX", "The keyword listing count disagrees with its records.");
            void Require(int bytes) { if (position > end - bytes) throw ChmBinary.Error("INDEX", "A keyword record is truncated."); }
        }
        return roots.Select(item => item.Freeze()).ToList().AsReadOnly();
    }

    private string ReadWideString(byte[] data, ref int position, int end) {
        int first = position;
        while (position <= end - 2) {
            if (position - first > (long)_options.MaxPathLength * 2) throw ChmBinary.Error("INDEX", "A keyword string exceeds MaxPathLength.");
            ushort value = ChmBinary.U16(data, position); position += 2;
            if (value == 0) return Encoding.Unicode.GetString(data, first, position - first - 2);
        }
        throw ChmBinary.Error("INDEX", "An index UTF-16 string is unterminated.");
    }
}
