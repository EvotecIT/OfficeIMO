using CffIndex = OfficeIMO.Pdf.PdfCffCharStringSubsetter.CffIndex;
using CffObject = OfficeIMO.Pdf.PdfCffCharStringSubsetter.CffObject;

namespace OfficeIMO.Pdf;

/// <summary>
/// Builds a compact CID-keyed CFF program (<c>/FontFile3 /Subtype /CIDFontType0C</c>) holding only the used glyphs.
/// </summary>
/// <remarks>
/// CIDs are the source glyph identifiers, so content-stream codes, <c>/W</c> and ToUnicode stay unchanged; the
/// charset maps each CID to its compact glyph. Subroutine numbering (and therefore bias) is preserved: unused
/// local and global subroutines become a bare <c>return</c>. Programs the analysis cannot prove complete
/// (seac-style accents, computed subroutine indexes, synthetic fonts, non-Type 2 charstrings) are declined,
/// keeping the OpenType embedding path.
/// </remarks>
internal static class PdfCffCidSubsetter {
    private const int TopDictOperatorCharset = 15;
    private const int TopDictOperatorEncoding = 16;
    private const int TopDictOperatorCharStrings = 17;
    private const int DictOperatorPrivate = 18;
    private const int PrivateDictOperatorSubrs = 19;
    private const int TopDictOperatorCharstringType = 1206;
    private const int TopDictOperatorSyntheticBase = 1220;
    private const int TopDictOperatorRos = 1230;
    private const int TopDictOperatorCidCount = 1234;
    private const int TopDictOperatorFdArray = 1236;
    private const int TopDictOperatorFdSelect = 1237;
    private const int StandardStringCount = 391;
    private const int MaximumSubroutineDepth = 10;
    private const int AnalysisOperationBudget = 20_000_000;
    private static readonly byte[] ReturnProgram = { 11 };
    private static readonly byte[] Header = { 1, 0, 4, 4 };
    private static readonly int[] NotDefGlyph = { 0 };

    internal static PdfCffCharStringSubset? TryCreate(byte[] cff, IReadOnlyList<int> usedGlyphIds, int expectedGlyphCount) {
        Guard.NotNull(cff, nameof(cff));
        Guard.NotNull(usedGlyphIds, nameof(usedGlyphIds));
        try {
            return Create(cff, usedGlyphIds, expectedGlyphCount);
        } catch (Exception exception) when (PdfCffCharStringSubsetter.IsCffDataException(exception)) {
            return null;
        }
    }

    private static PdfCffCharStringSubset? Create(byte[] cff, IReadOnlyList<int> usedGlyphIds, int expectedGlyphCount) {
        if (cff.Length < 4 || cff[0] != 1) return null;
        int headerSize = cff[2];
        if (headerSize < 4 || headerSize > cff.Length) return null;
        CffIndex names = PdfCffCharStringSubsetter.ReadIndex(cff, headerSize);
        CffIndex topDictionaries = PdfCffCharStringSubsetter.ReadIndex(cff, names.EndOffset);
        if (topDictionaries.Count != 1) return null;
        CffIndex strings = PdfCffCharStringSubsetter.ReadIndex(cff, topDictionaries.EndOffset);
        CffIndex globalSubrs = PdfCffCharStringSubsetter.ReadIndex(cff, strings.EndOffset);
        List<DictEntry> top = ParseDict(cff, topDictionaries.GetObject(0));
        if (Find(top, TopDictOperatorSyntheticBase) != null) return null;
        DictEntry? charstringType = Find(top, TopDictOperatorCharstringType);
        if (charstringType != null && charstringType.Value.Integer(0) != 2) return null;
        DictEntry? charStringsEntry = Find(top, TopDictOperatorCharStrings);
        if (charStringsEntry == null) return null;
        CffIndex charStrings = PdfCffCharStringSubsetter.ReadIndex(cff, charStringsEntry.Value.Integer(0));
        if (charStrings.Count != expectedGlyphCount || charStrings.Count == 0) return null;

        int[] glyphs = usedGlyphIds.Concat(NotDefGlyph).Distinct().OrderBy(static id => id).ToArray();
        if (glyphs[glyphs.Length - 1] >= charStrings.Count || glyphs[0] < 0) return null;
        if (glyphs.Length >= charStrings.Count) return null;

        // Font dictionaries: name-keyed sources have one implicit dictionary (the Top DICT Private).
        bool cidKeyed = Find(top, TopDictOperatorRos) != null;
        var fontDictionaries = new List<FontDictionary>();
        int[] fdOfGlyph;
        if (cidKeyed) {
            DictEntry? fdArrayEntry = Find(top, TopDictOperatorFdArray);
            DictEntry? fdSelectEntry = Find(top, TopDictOperatorFdSelect);
            if (fdArrayEntry == null || fdSelectEntry == null) return null;
            CffIndex fdArray = PdfCffCharStringSubsetter.ReadIndex(cff, fdArrayEntry.Value.Integer(0));
            fdOfGlyph = ReadFdSelect(cff, fdSelectEntry.Value.Integer(0), charStrings.Count, fdArray.Count);
            var usedFds = new bool[fdArray.Count];
            foreach (int glyph in glyphs) usedFds[fdOfGlyph[glyph]] = true;
            FontDictionary? unusedDictionary = null;
            for (int index = 0; index < fdArray.Count; index++) {
                if (usedFds[index]) {
                    List<DictEntry> entries = ParseDict(cff, fdArray.GetObject(index));
                    fontDictionaries.Add(ReadFontDictionary(cff, entries));
                } else {
                    unusedDictionary ??= new FontDictionary(new List<DictEntry>(), new List<DictEntry>(), EmptyIndex());
                    fontDictionaries.Add(unusedDictionary);
                }
            }
        } else {
            fontDictionaries.Add(ReadFontDictionary(cff, new List<DictEntry>(), Find(top, DictOperatorPrivate)));
            fdOfGlyph = new int[charStrings.Count];
        }

        var usedGlobal = new bool[globalSubrs.Count];
        var analyzer = new CharStringAnalyzer(cff, globalSubrs, usedGlobal);
        foreach (int glyph in glyphs) {
            FontDictionary fd = fontDictionaries[fdOfGlyph[glyph]];
            if (!analyzer.Analyze(charStrings.GetObject(glyph), fd.LocalSubrs, fd.UsedLocal)) return null;
        }

        // Retained font dictionaries keep their source order.
        int[] newFdOfOld = Enumerable.Repeat(-1, fontDictionaries.Count).ToArray();
        var keptFds = new List<int>();
        foreach (int glyph in glyphs) {
            int fd = fdOfGlyph[glyph];
            if (newFdOfOld[fd] < 0) newFdOfOld[fd] = -2;
        }
        for (int fd = 0; fd < fontDictionaries.Count; fd++) {
            if (newFdOfOld[fd] == -2) {
                newFdOfOld[fd] = keptFds.Count;
                keptFds.Add(fd);
            }
        }

        // Strings: glyph names are not needed by a CID-keyed font, so only dictionary strings are kept.
        var stringTable = new StringTable(cff, strings);
        int registrySid = stringTable.Add(System.Text.Encoding.ASCII.GetBytes("Adobe"));
        int orderingSid = stringTable.Add(System.Text.Encoding.ASCII.GetBytes("Identity"));

        byte[] nameIndex = Slice(cff, names.StartOffset, names.TotalLength);
        // Dictionaries are built (and their strings registered) before the String INDEX is laid out.
        BuildTopDict(cff, top, stringTable, registrySid, orderingSid, 0, 0, 0, 0, 0);
        BuildFdArray(cff, fontDictionaries, keptFds, stringTable, new List<byte[]>(keptFds.Select(static _ => Array.Empty<byte>())), new int[keptFds.Count]);
        byte[] stringIndex = BuildIndex(stringTable.Strings);
        byte[] globalSubrIndex = BuildSubrIndex(cff, globalSubrs, usedGlobal);
        byte[] charset = BuildCharset(glyphs);
        byte[] fdSelect = BuildFdSelect(glyphs.Select(glyph => newFdOfOld[fdOfGlyph[glyph]]).ToArray());
        var programs = new List<byte[]>(glyphs.Length);
        int subsetProgramBytes = 0;
        foreach (int glyph in glyphs) {
            byte[] program = PdfCffCharStringSubsetter.CopyObject(cff, charStrings.GetObject(glyph));
            programs.Add(program);
            subsetProgramBytes = checked(subsetProgramBytes + program.Length);
        }
        byte[] charStringIndex = BuildIndex(programs);

        var privates = new List<byte[]>(keptFds.Count);
        var localSubrIndexes = new List<byte[]?>(keptFds.Count);
        foreach (int fd in keptFds) {
            FontDictionary dictionary = fontDictionaries[fd];
            byte[]? subrs = dictionary.LocalSubrs.Count > 0 ? BuildSubrIndex(cff, dictionary.LocalSubrs, dictionary.UsedLocal) : null;
            var privateDict = new List<byte>();
            foreach (DictEntry entry in dictionary.PrivateEntries) {
                if (entry.Operator != PrivateDictOperatorSubrs) privateDict.AddRange(Slice(cff, entry.Start, entry.End - entry.Start));
            }
            if (subrs != null) {
                int privateSize = privateDict.Count + 6;
                AppendInteger(privateDict, privateSize);
                AppendOperator(privateDict, PrivateDictOperatorSubrs);
            }
            privates.Add(privateDict.ToArray());
            localSubrIndexes.Add(subrs);
        }

        // Fixed-width offsets make the Top DICT and FDArray sizes independent of the layout they describe.
        byte[] topDict = BuildTopDict(cff, top, stringTable, registrySid, orderingSid, glyphs[glyphs.Length - 1] + 1, 0, 0, 0, 0);
        byte[] topIndex = BuildIndex(new List<byte[]> { topDict });
        int charsetOffset = 4 + nameIndex.Length + topIndex.Length + stringIndex.Length + globalSubrIndex.Length;
        int fdSelectOffset = charsetOffset + charset.Length;
        int charStringsOffset = fdSelectOffset + fdSelect.Length;
        int fdArrayOffset = charStringsOffset + charStringIndex.Length;
        byte[] fdArrayIndex = BuildFdArray(cff, fontDictionaries, keptFds, stringTable, privates, new int[keptFds.Count]);
        var privateOffsets = new int[keptFds.Count];
        int cursor = fdArrayOffset + fdArrayIndex.Length;
        for (int index = 0; index < keptFds.Count; index++) {
            privateOffsets[index] = cursor;
            cursor = checked(cursor + privates[index].Length + (localSubrIndexes[index]?.Length ?? 0));
        }
        fdArrayIndex = BuildFdArray(cff, fontDictionaries, keptFds, stringTable, privates, privateOffsets);
        topDict = BuildTopDict(cff, top, stringTable, registrySid, orderingSid, glyphs[glyphs.Length - 1] + 1, charsetOffset, charStringsOffset, fdArrayOffset, fdSelectOffset);
        topIndex = BuildIndex(new List<byte[]> { topDict });

        using var output = new MemoryStream(cursor);
        output.Write(Header, 0, Header.Length);
        Write(output, nameIndex);
        Write(output, topIndex);
        Write(output, stringIndex);
        Write(output, globalSubrIndex);
        Write(output, charset);
        Write(output, fdSelect);
        Write(output, charStringIndex);
        Write(output, fdArrayIndex);
        for (int index = 0; index < keptFds.Count; index++) {
            Write(output, privates[index]);
            if (localSubrIndexes[index] != null) Write(output, localSubrIndexes[index]!);
        }
        if (output.Length != cursor) throw new NotSupportedException("The compact CFF layout is inconsistent.");
        byte[] data = output.ToArray();
        return PdfCffCharStringSubset.Subset(data, charStrings.Count, glyphs.Length, charStrings.Count - glyphs.Length,
            charStrings.DataLength, subsetProgramBytes, charStrings.TotalLength, charStringIndex.Length);
    }

    private static FontDictionary ReadFontDictionary(byte[] cff, List<DictEntry> entries, DictEntry? privateEntry = null) {
        privateEntry ??= Find(entries, DictOperatorPrivate);
        if (privateEntry == null) return new FontDictionary(entries, new List<DictEntry>(), EmptyIndex());
        int size = privateEntry.Value.Integer(0);
        int offset = privateEntry.Value.Integer(1);
        if (size < 0 || offset < 0 || offset > cff.Length - size) throw new NotSupportedException("The CFF Private DICT range is invalid.");
        List<DictEntry> privateEntries = ParseDict(cff, new CffObject(offset, size));
        DictEntry? subrs = Find(privateEntries, PrivateDictOperatorSubrs);
        CffIndex localSubrs = subrs == null ? EmptyIndex() : PdfCffCharStringSubsetter.ReadIndex(cff, checked(offset + subrs.Value.Integer(0)));
        return new FontDictionary(entries, privateEntries, localSubrs);
    }

    private static CffIndex EmptyIndex() => new CffIndex(0, 0, 0, 0, Array.Empty<int>());

    private static int[] ReadFdSelect(byte[] cff, int offset, int glyphCount, int fdCount) {
        var result = new int[glyphCount];
        int format = cff[offset];
        if (format == 0) {
            for (int glyph = 0; glyph < glyphCount; glyph++) result[glyph] = cff[checked(offset + 1 + glyph)];
        } else if (format == 3) {
            int ranges = ReadUInt16(cff, offset + 1);
            int position = offset + 3;
            for (int range = 0; range < ranges; range++) {
                int first = ReadUInt16(cff, position);
                int fd = cff[position + 2];
                int next = ReadUInt16(cff, position + 3);
                if (first > next || next > glyphCount) throw new NotSupportedException("The CFF FDSelect range is invalid.");
                for (int glyph = first; glyph < next; glyph++) result[glyph] = fd;
                position += 3;
            }
        } else {
            throw new NotSupportedException("The CFF FDSelect format is not supported.");
        }
        foreach (int fd in result) {
            if (fd >= fdCount) throw new NotSupportedException("The CFF FDSelect references a missing font dictionary.");
        }
        return result;
    }

    private static byte[] BuildTopDict(byte[] cff, List<DictEntry> top, StringTable strings, int registrySid, int orderingSid, int cidCount,
        int charset, int charStrings, int fdArray, int fdSelect) {
        var result = new List<byte>();
        // ROS must be the first Top DICT operator of a CID-keyed font.
        AppendInteger(result, registrySid);
        AppendInteger(result, orderingSid);
        AppendInteger(result, 0);
        AppendOperator(result, TopDictOperatorRos);
        foreach (DictEntry entry in top) {
            if (entry.Operator is TopDictOperatorCharset or TopDictOperatorEncoding or TopDictOperatorCharStrings or DictOperatorPrivate ||
                entry.Operator is >= TopDictOperatorRos and <= 1238 ||
                entry.Operator is 14 or 1221 or 1222 or 1223) continue; // XUID, PostScript, BaseFontName, BaseFontBlend
            AppendEntry(result, cff, entry, strings);
        }
        AppendInteger(result, cidCount);
        AppendOperator(result, TopDictOperatorCidCount);
        AppendInteger(result, charset);
        AppendOperator(result, TopDictOperatorCharset);
        AppendInteger(result, charStrings);
        AppendOperator(result, TopDictOperatorCharStrings);
        AppendInteger(result, fdArray);
        AppendOperator(result, TopDictOperatorFdArray);
        AppendInteger(result, fdSelect);
        AppendOperator(result, TopDictOperatorFdSelect);
        return result.ToArray();
    }

    private static byte[] BuildFdArray(byte[] cff, List<FontDictionary> dictionaries, List<int> keptFds, StringTable strings, List<byte[]> privates, int[] privateOffsets) {
        var objects = new List<byte[]>(keptFds.Count);
        for (int index = 0; index < keptFds.Count; index++) {
            var result = new List<byte>();
            foreach (DictEntry entry in dictionaries[keptFds[index]].Entries) {
                if (entry.Operator != DictOperatorPrivate) AppendEntry(result, cff, entry, strings);
            }
            AppendInteger(result, privates[index].Length);
            AppendInteger(result, privateOffsets[index]);
            AppendOperator(result, DictOperatorPrivate);
            objects.Add(result.ToArray());
        }
        return BuildIndex(objects);
    }

    // version, Notice, FullName, FamilyName, Weight, Copyright and FontName carry one SID operand.
    private static void AppendEntry(List<byte> output, byte[] cff, DictEntry entry, StringTable strings) {
        if (entry.Operator is 0 or 1 or 2 or 3 or 4 or 1200 or 1238 && entry.Operands.Length == 1) {
            AppendInteger(output, strings.Map(entry.Integer(0)));
            AppendOperator(output, entry.Operator);
            return;
        }
        output.AddRange(Slice(cff, entry.Start, entry.End - entry.Start));
    }

    private static byte[] BuildCharset(int[] glyphs) {
        // Format 0 lists every CID; format 2 lists runs of consecutive CIDs. Glyph 0 is implicit.
        int ranges = 0;
        for (int index = 1; index < glyphs.Length; index++) {
            if (index == 1 || glyphs[index] != glyphs[index - 1] + 1) ranges++;
        }
        var result = new List<byte>();
        if (1 + 2 * (glyphs.Length - 1) <= 1 + 4 * ranges) {
            result.Add(0);
            for (int index = 1; index < glyphs.Length; index++) AppendUInt16(result, glyphs[index]);
            return result.ToArray();
        }
        result.Add(2);
        int start = 1;
        while (start < glyphs.Length) {
            int end = start;
            while (end + 1 < glyphs.Length && glyphs[end + 1] == glyphs[end] + 1 && end - start < ushort.MaxValue) end++;
            AppendUInt16(result, glyphs[start]);
            AppendUInt16(result, end - start);
            start = end + 1;
        }
        return result.ToArray();
    }

    private static byte[] BuildFdSelect(int[] fdOfNewGlyph) {
        var ranges = new List<(int First, int Fd)>();
        for (int glyph = 0; glyph < fdOfNewGlyph.Length; glyph++) {
            if (glyph == 0 || fdOfNewGlyph[glyph] != fdOfNewGlyph[glyph - 1]) ranges.Add((glyph, fdOfNewGlyph[glyph]));
        }
        var result = new List<byte> { 3 };
        AppendUInt16(result, ranges.Count);
        foreach ((int first, int fd) in ranges) {
            AppendUInt16(result, first);
            result.Add(checked((byte)fd));
        }
        AppendUInt16(result, fdOfNewGlyph.Length);
        return result.ToArray();
    }

    private static byte[] BuildSubrIndex(byte[] cff, CffIndex subrs, bool[] used) {
        var objects = new List<byte[]>(subrs.Count);
        for (int index = 0; index < subrs.Count; index++) {
            objects.Add(used[index] ? PdfCffCharStringSubsetter.CopyObject(cff, subrs.GetObject(index)) : ReturnProgram);
        }
        return BuildIndex(objects);
    }

    private static byte[] BuildIndex(List<byte[]> objects) {
        if (objects.Count == 0) return new byte[] { 0, 0 };
        int length = 0;
        foreach (byte[] value in objects) length = checked(length + value.Length);
        return PdfCffCharStringSubsetter.BuildIndex(objects.ToArray(), length);
    }

    private static List<DictEntry> ParseDict(byte[] data, CffObject dictionary) {
        var entries = new List<DictEntry>();
        var operands = new List<double>();
        int offset = dictionary.Offset;
        int end = checked(dictionary.Offset + dictionary.Length);
        if (offset < 0 || end > data.Length) throw new NotSupportedException("The CFF DICT range is invalid.");
        int entryStart = offset;
        while (offset < end) {
            int value = data[offset++];
            if (value >= 32 || value == 28 || value == 29 || value == 30 || value == 255) {
                operands.Add(PdfCffCharStringSubsetter.ReadDictionaryNumber(data, ref offset, end, value));
                continue;
            }
            int dictionaryOperator = value;
            if (value == 12) {
                if (offset >= end) throw new NotSupportedException("The CFF DICT contains a truncated escaped operator.");
                dictionaryOperator = 1200 + data[offset++];
            }
            entries.Add(new DictEntry(dictionaryOperator, entryStart, offset, operands.ToArray()));
            operands.Clear();
            entryStart = offset;
        }
        return entries;
    }

    private static DictEntry? Find(List<DictEntry> entries, int dictionaryOperator) {
        foreach (DictEntry entry in entries) {
            if (entry.Operator == dictionaryOperator) return entry;
        }
        return null;
    }

    private static void AppendInteger(List<byte> output, int value) {
        output.Add(29);
        output.Add((byte)(value >> 24));
        output.Add((byte)(value >> 16));
        output.Add((byte)(value >> 8));
        output.Add((byte)value);
    }

    private static void AppendOperator(List<byte> output, int dictionaryOperator) {
        if (dictionaryOperator >= 1200) {
            output.Add(12);
            output.Add((byte)(dictionaryOperator - 1200));
        } else {
            output.Add((byte)dictionaryOperator);
        }
    }

    private static void AppendUInt16(List<byte> output, int value) {
        if (value < 0 || value > ushort.MaxValue) throw new NotSupportedException("A CFF Card16 value is out of range.");
        output.Add((byte)(value >> 8));
        output.Add((byte)value);
    }

    private static int ReadUInt16(byte[] data, int offset) => (data[offset] << 8) | data[offset + 1];

    private static byte[] Slice(byte[] data, int offset, int length) {
        var result = new byte[length];
        Buffer.BlockCopy(data, offset, result, 0, length);
        return result;
    }

    private static void Write(Stream output, byte[] value) => output.Write(value, 0, value.Length);

    private readonly struct DictEntry {
        internal DictEntry(int dictionaryOperator, int start, int end, double[] operands) {
            Operator = dictionaryOperator;
            Start = start;
            End = end;
            Operands = operands;
        }

        internal int Operator { get; }
        internal int Start { get; }
        internal int End { get; }
        internal double[] Operands { get; }

        internal int Integer(int index) {
            if (index >= Operands.Length) throw new NotSupportedException("A CFF DICT operator is missing an operand.");
            double value = Operands[index];
            if (value < int.MinValue || value > int.MaxValue || Math.Abs(value - Math.Round(value)) > 0.00001D) {
                throw new NotSupportedException("A CFF DICT operand is not an integer.");
            }
            return (int)Math.Round(value);
        }
    }

    /// <summary>Custom strings referenced by the retained dictionaries, renumbered densely after the standard strings.</summary>
    private sealed class StringTable {
        private readonly byte[] _cff;
        private readonly CffIndex _source;
        private readonly Dictionary<int, int> _mapped = new Dictionary<int, int>();

        internal StringTable(byte[] cff, CffIndex source) {
            _cff = cff;
            _source = source;
        }

        internal List<byte[]> Strings { get; } = new List<byte[]>();

        internal int Add(byte[] value) {
            Strings.Add(value);
            return StandardStringCount + Strings.Count - 1;
        }

        internal int Map(int sid) {
            if (sid < StandardStringCount) return sid;
            if (_mapped.TryGetValue(sid, out int mapped)) return mapped;
            mapped = Add(PdfCffCharStringSubsetter.CopyObject(_cff, _source.GetObject(sid - StandardStringCount)));
            _mapped[sid] = mapped;
            return mapped;
        }
    }

    private sealed class FontDictionary {
        internal FontDictionary(List<DictEntry> entries, List<DictEntry> privateEntries, CffIndex localSubrs) {
            Entries = entries;
            PrivateEntries = privateEntries;
            LocalSubrs = localSubrs;
            UsedLocal = new bool[localSubrs.Count];
        }

        internal List<DictEntry> Entries { get; }
        internal List<DictEntry> PrivateEntries { get; }
        internal CffIndex LocalSubrs { get; }
        internal bool[] UsedLocal { get; }
    }

    /// <summary>Walks Type 2 charstrings to find every reachable subroutine, tracking stem hints for mask lengths.</summary>
    private sealed class CharStringAnalyzer {
        private readonly byte[] _data;
        private readonly CffIndex _globalSubrs;
        private readonly bool[] _usedGlobal;
        private readonly List<double> _stack = new List<double>(48);
        private int _stems;
        private int _operations;
        private bool _ended;

        internal CharStringAnalyzer(byte[] data, CffIndex globalSubrs, bool[] usedGlobal) {
            _data = data;
            _globalSubrs = globalSubrs;
            _usedGlobal = usedGlobal;
        }

        internal bool Analyze(CffObject program, CffIndex localSubrs, bool[] usedLocal) {
            _stack.Clear();
            _stems = 0;
            _ended = false;
            return Execute(program, localSubrs, usedLocal, 0);
        }

        private bool Execute(CffObject program, CffIndex localSubrs, bool[] usedLocal, int depth) {
            int position = program.Offset;
            int end = program.Offset + program.Length;
            while (position < end) {
                if (++_operations > AnalysisOperationBudget) return false;
                int value = _data[position++];
                if (value == 28) {
                    if (position + 2 > end) return false;
                    _stack.Add((short)((_data[position] << 8) | _data[position + 1]));
                    position += 2;
                } else if (value >= 32 && value <= 246) {
                    _stack.Add(value - 139);
                } else if (value >= 247 && value <= 250) {
                    if (position >= end) return false;
                    _stack.Add((value - 247) * 256 + _data[position++] + 108);
                } else if (value >= 251 && value <= 254) {
                    if (position >= end) return false;
                    _stack.Add(-(value - 251) * 256 - _data[position++] - 108);
                } else if (value == 255) {
                    if (position + 4 > end) return false;
                    int fixedValue = (_data[position] << 24) | (_data[position + 1] << 16) | (_data[position + 2] << 8) | _data[position + 3];
                    _stack.Add(fixedValue / 65536D);
                    position += 4;
                } else if (value is 1 or 3 or 18 or 23) {
                    _stems += _stack.Count / 2;
                    _stack.Clear();
                } else if (value is 19 or 20) {
                    _stems += _stack.Count / 2;
                    _stack.Clear();
                    position += (_stems + 7) / 8;
                    if (position > end) return false;
                } else if (value is 10 or 29) {
                    if (_stack.Count == 0 || depth >= MaximumSubroutineDepth) return false;
                    double operand = _stack[_stack.Count - 1];
                    _stack.RemoveAt(_stack.Count - 1);
                    CffIndex subrs = value == 10 ? localSubrs : _globalSubrs;
                    bool[] used = value == 10 ? usedLocal : _usedGlobal;
                    if (Math.Abs(operand - Math.Round(operand)) > 0.00001D) return false;
                    int index = (int)Math.Round(operand) + Bias(subrs.Count);
                    if (index < 0 || index >= subrs.Count) return false;
                    used[index] = true;
                    if (!Execute(subrs.GetObject(index), localSubrs, usedLocal, depth + 1)) return false;
                    if (_ended) return true;
                } else if (value == 11) {
                    return true;
                } else if (value == 14) {
                    // endchar with four (or five, including width) arguments is the deprecated seac accent form.
                    if (_stack.Count >= 4) return false;
                    _ended = true;
                    return true;
                } else if (value == 12) {
                    if (position >= end) return false;
                    int escaped = _data[position++];
                    // Arithmetic, storage and conditional operators could compute subroutine indexes.
                    if (escaped is not (0 or 34 or 35 or 36 or 37)) return false;
                    _stack.Clear();
                } else if (value is 4 or 5 or 6 or 7 or 8 or 21 or 22 or 24 or 25 or 26 or 27 or 30 or 31) {
                    _stack.Clear();
                } else {
                    return false;
                }
            }
            return true;
        }

        private static int Bias(int count) => count < 1240 ? 107 : count < 33900 ? 1131 : 32768;
    }
}
