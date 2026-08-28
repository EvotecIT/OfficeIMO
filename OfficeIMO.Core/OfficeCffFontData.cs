using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>Validated CFF1/CFF2 table model used by the first-party OpenType outline program.</summary>
internal sealed class OfficeCffFontData {
    private const int MaximumIndexObjects = 1_000_000;
    private readonly CffIndex _charStrings;
    private readonly CffIndex _globalSubroutines;
    private readonly CffIndex[] _localSubroutines;
    private readonly int[]? _fontDictionaryByGlyph;
    private readonly int[]? _standardEncodingGlyphs;

    private OfficeCffFontData(
        bool isCff2,
        int maximumOperandStack,
        CffIndex charStrings,
        CffIndex globalSubroutines,
        CffIndex[] localSubroutines,
        int[]? fontDictionaryByGlyph,
        int[]? standardEncodingGlyphs,
        OfficeCffVariationStore? variationStore) {
        IsCff2 = isCff2;
        MaximumOperandStack = maximumOperandStack;
        _charStrings = charStrings;
        _globalSubroutines = globalSubroutines;
        _localSubroutines = localSubroutines;
        _fontDictionaryByGlyph = fontDictionaryByGlyph;
        _standardEncodingGlyphs = standardEncodingGlyphs;
        VariationStore = variationStore;
    }

    internal bool IsCff2 { get; }
    internal int MaximumOperandStack { get; }
    internal int GlyphCount => _charStrings.Count;
    internal OfficeCffVariationStore? VariationStore { get; }

    internal static bool IsStructurallyValidProgram(byte[] data, bool isCff2, bool? requireCidKeyed = null) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        try {
            int minimumHeaderSize = isCff2 ? 5 : 4;
            if (data.Length < minimumHeaderSize ||
                data[0] != (isCff2 ? 2 : 1) ||
                data[2] < minimumHeaderSize ||
                data[2] > data.Length) {
                return false;
            }

            int tableEnd = data.Length;
            if (isCff2) {
                int topLength = ReadUInt16(data, 3, tableEnd);
                int topOffset = data[2];
                if (topLength <= 0) return false;
                EnsureRange(topOffset, topLength, 0, tableEnd, "The CFF2 Top DICT is truncated.");
                CffDictionary topDictionary = CffDictionary.Parse(data, topOffset, topLength);
                int cursor = checked(topOffset + topLength);
                CffIndex cff2GlobalSubroutines = CffIndex.Read(data, ref cursor, tableEnd, countSize: 4);
                CffIndex cff2CharStrings = ReadNonEmptyCharStrings(data, tableEnd, topDictionary, cursor, countSize: 4);
                CffIndex localSubroutines = topDictionary.TryGetPair(18, out int privateSize, out int privateOffset)
                    ? ReadLocalSubroutines(data, 0, tableEnd, privateSize, privateOffset, isCff2: true)
                    : CffIndex.Empty(data);
                ValidateCharStrings(
                    isCff2: true,
                    cff2CharStrings,
                    cff2GlobalSubroutines,
                    new[] { localSubroutines },
                    fontDictionaryByGlyph: null,
                    standardEncodingGlyphs: null);
                return true;
            }

            if (data[3] < 1 || data[3] > 4) return false;
            int cff1Cursor = data[2];
            CffIndex names = CffIndex.Read(data, ref cff1Cursor, tableEnd, countSize: 2);
            CffIndex topDictionaries = CffIndex.Read(data, ref cff1Cursor, tableEnd, countSize: 2);
            if (names.Count != 1 || topDictionaries.Count != 1) return false;
            _ = CffIndex.Read(data, ref cff1Cursor, tableEnd, countSize: 2);
            CffIndex globalSubroutines = CffIndex.Read(data, ref cff1Cursor, tableEnd, countSize: 2);
            CffSlice top = topDictionaries[0];
            CffDictionary cff1TopDictionary = CffDictionary.Parse(data, top.Offset, top.Length);
            bool isCidKeyed = cff1TopDictionary.ContainsOperation(0x0C1E);
            if (requireCidKeyed.HasValue &&
                isCidKeyed != requireCidKeyed.Value) return false;
            CffIndex charStrings = ReadNonEmptyCharStrings(
                data,
                tableEnd,
                cff1TopDictionary,
                cff1Cursor,
                countSize: 2);
            if (!isCidKeyed) {
                CffIndex localSubroutines = cff1TopDictionary.TryGetPair(18, out int privateSize, out int privateOffset)
                    ? ReadLocalSubroutines(data, 0, tableEnd, privateSize, privateOffset, isCff2: false)
                    : CffIndex.Empty(data);
                int charset = cff1TopDictionary.TryGetInteger(15, out int selectedCharset) ? selectedCharset : 0;
                int[] standardEncodingGlyphs = OfficeCffCharset.BuildStandardEncodingGlyphMap(
                    data,
                    0,
                    tableEnd,
                    charset,
                    charStrings.Count);
                ValidateCharStrings(
                    isCff2: false,
                    charStrings,
                    globalSubroutines,
                    new[] { localSubroutines },
                    fontDictionaryByGlyph: null,
                    standardEncodingGlyphs);
                return true;
            }
            if (!cff1TopDictionary.HasNonNegativeIntegerOperands(0x0C1E, 3) ||
                !cff1TopDictionary.TryGetInteger(0x0C24, out int fdArrayOffset) ||
                !cff1TopDictionary.TryGetInteger(0x0C25, out int fdSelectOffset) ||
                fdArrayOffset < cff1Cursor || fdSelectOffset < cff1Cursor) return false;
            CffIndex fdArray = CffIndex.ReadAt(data, fdArrayOffset, tableEnd, countSize: 2);
            if (fdArray.Count <= 0 || fdArray.Count > 256) return false;
            var localSubroutinesByDictionary = new CffIndex[fdArray.Count];
            for (int index = 0; index < fdArray.Count; index++) {
                CffSlice fontDictionary = fdArray[index];
                CffDictionary dictionary = CffDictionary.Parse(data, fontDictionary.Offset, fontDictionary.Length);
                localSubroutinesByDictionary[index] = dictionary.TryGetPair(18, out int privateSize, out int privateOffset)
                    ? ReadLocalSubroutines(data, 0, tableEnd, privateSize, privateOffset, isCff2: false)
                    : CffIndex.Empty(data);
            }
            int[] fontDictionaryByGlyph = ReadFdSelect(data, fdSelectOffset, tableEnd, charStrings.Count, fdArray.Count);
            ValidateCharStrings(
                isCff2: false,
                charStrings,
                globalSubroutines,
                localSubroutinesByDictionary,
                fontDictionaryByGlyph,
                standardEncodingGlyphs: null);
            return true;
        } catch (Exception exception) when (
            exception is InvalidDataException ||
            exception is OverflowException ||
            exception is ArgumentOutOfRangeException ||
            exception is NotSupportedException) {
            return false;
        }
    }

    private static void ValidateCharStrings(
        bool isCff2,
        CffIndex charStrings,
        CffIndex globalSubroutines,
        CffIndex[] localSubroutines,
        int[]? fontDictionaryByGlyph,
        int[]? standardEncodingGlyphs) {
        var font = new OfficeCffFontData(
            isCff2,
            isCff2 ? 513 : 48,
            charStrings,
            globalSubroutines,
            localSubroutines,
            fontDictionaryByGlyph,
            standardEncodingGlyphs,
            variationStore: null);
        var operationBudget = new OfficeCffOperationBudget();
        for (int glyphId = 0; glyphId < charStrings.Count; glyphId++) {
            new OfficeType2CharStringInterpreter(
                font,
                glyphId,
                StructuralValidationSink.Instance,
                System.Threading.CancellationToken.None,
                operationBudget).Render(charStrings[glyphId]);
        }
    }

    private sealed class StructuralValidationSink : IOfficeCffPathSink {
        internal static StructuralValidationSink Instance { get; } = new();
        public void MoveTo(double x, double y) { }
        public void LineTo(double x, double y) { }
        public void CurveTo(double control1X, double control1Y, double control2X, double control2Y, double x, double y) { }
        public void CloseContour() { }
    }

    private static CffIndex ReadNonEmptyCharStrings(
        byte[] data,
        int tableEnd,
        CffDictionary topDictionary,
        int minimumOffset,
        int countSize) {
        int charStringsOffset = topDictionary.GetRequiredInteger(17, "CharStrings");
        if (charStringsOffset < minimumOffset) {
            throw new InvalidDataException("The CFF CharStrings INDEX overlaps preceding data.");
        }
        EnsureRange(charStringsOffset, countSize, 0, tableEnd, "The CFF CharStrings INDEX offset is invalid.");
        CffIndex charStrings = CffIndex.ReadAt(data, charStringsOffset, tableEnd, countSize);
        if (charStrings.Count <= 0) throw new InvalidDataException("The CFF CharStrings INDEX is empty.");
        return charStrings;
    }

    internal static OfficeCffFontData Parse(OfficeOpenTypeReader reader, OfficeFontVariationModel variations) {
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        if (variations == null) throw new ArgumentNullException(nameof(variations));
        bool isCff2 = reader.TryGetTable("CFF2", out int tableOffset, out int tableLength);
        if (!isCff2 && !reader.TryGetTable("CFF ", out tableOffset, out tableLength)) {
            throw new InvalidDataException("The OpenType font does not contain CFF outlines.");
        }
        byte[] data = reader.Data;
        int tableEnd = checked(tableOffset + tableLength);
        if (tableLength < (isCff2 ? 5 : 4)) throw new InvalidDataException("The CFF table header is truncated.");
        int major = data[tableOffset];
        int headerSize = data[tableOffset + 2];
        if ((isCff2 && major != 2) || (!isCff2 && major != 1)
            || headerSize < (isCff2 ? 5 : 4) || headerSize > tableLength) {
            throw new InvalidDataException("The CFF table header is invalid.");
        }

        CffDictionary topDictionary;
        CffIndex globalSubroutines;
        if (isCff2) {
            int topLength = ReadUInt16(data, tableOffset + 3, tableEnd);
            int topOffset = checked(tableOffset + headerSize);
            EnsureRange(topOffset, topLength, tableOffset, tableEnd, "The CFF2 Top DICT is truncated.");
            topDictionary = CffDictionary.Parse(data, topOffset, topLength);
            int cursor = checked(topOffset + topLength);
            globalSubroutines = CffIndex.Read(data, ref cursor, tableEnd, countSize: 4);
        } else {
            int cursor = checked(tableOffset + headerSize);
            CffIndex names = CffIndex.Read(data, ref cursor, tableEnd, countSize: 2);
            if (names.Count != 1) throw new NotSupportedException("Only single-font CFF1 tables are supported.");
            CffIndex topDictionaries = CffIndex.Read(data, ref cursor, tableEnd, countSize: 2);
            if (topDictionaries.Count != 1) throw new InvalidDataException("The CFF1 Top DICT INDEX is invalid.");
            CffSlice top = topDictionaries[0];
            topDictionary = CffDictionary.Parse(data, top.Offset, top.Length);
            _ = CffIndex.Read(data, ref cursor, tableEnd, countSize: 2); // String INDEX
            globalSubroutines = CffIndex.Read(data, ref cursor, tableEnd, countSize: 2);
        }

        int charStringsRelative = topDictionary.GetRequiredInteger(17, "CharStrings");
        int charStringsOffset = checked(tableOffset + charStringsRelative);
        EnsureRange(charStringsOffset, 1, tableOffset, tableEnd, "The CFF CharStrings INDEX offset is invalid.");
        CffIndex charStrings = CffIndex.ReadAt(data, charStringsOffset, tableEnd, isCff2 ? 4 : 2);
        if (charStrings.Count <= 0 || charStrings.Count != reader.GlyphCount) {
            throw new InvalidDataException("The CFF CharStrings count does not match the OpenType glyph count.");
        }

        OfficeCffVariationStore? variationStore = null;
        if (isCff2 && topDictionary.TryGetInteger(24, out int variationStoreRelative)) {
            variationStore = OfficeCffVariationStore.Parse(
                reader,
                checked(tableOffset + variationStoreRelative),
                tableEnd,
                variations);
        }
        int maximumOperandStack = isCff2 && topDictionary.TryGetInteger(25, out int declaredMaximumStack)
            ? declaredMaximumStack
            : isCff2 ? 513 : 48;
        if (maximumOperandStack <= 0 || maximumOperandStack > 513) {
            throw new InvalidDataException("The CFF operand-stack limit is invalid.");
        }

        int[]? standardEncodingGlyphs = null;
        if (!isCff2) {
            int charsetValue = topDictionary.TryGetInteger(15, out int selectedCharset) ? selectedCharset : 0;
            if (charsetValue < 0) throw new InvalidDataException("The CFF charset offset is invalid.");
            standardEncodingGlyphs = OfficeCffCharset.BuildStandardEncodingGlyphMap(
                data,
                tableOffset,
                tableEnd,
                charsetValue,
                charStrings.Count);
        }

        CffIndex[] localSubroutines;
        int[]? fontDictionaryByGlyph = null;
        if (topDictionary.TryGetPair(18, out int privateSize, out int privateRelative)) {
            localSubroutines = new[] {
                ReadLocalSubroutines(data, tableOffset, tableEnd, privateSize, privateRelative, isCff2)
            };
        } else if (topDictionary.TryGetInteger(0x0C24, out int fdArrayRelative)) {
            int fdArrayOffset = checked(tableOffset + fdArrayRelative);
            CffIndex fdArray = CffIndex.ReadAt(data, fdArrayOffset, tableEnd, isCff2 ? 4 : 2);
            if (fdArray.Count <= 0 || fdArray.Count > 256) throw new InvalidDataException("The CFF FDArray is invalid.");
            localSubroutines = new CffIndex[fdArray.Count];
            for (int index = 0; index < fdArray.Count; index++) {
                CffSlice dictionaryBytes = fdArray[index];
                CffDictionary dictionary = CffDictionary.Parse(data, dictionaryBytes.Offset, dictionaryBytes.Length);
                localSubroutines[index] = dictionary.TryGetPair(18, out int size, out int relative)
                    ? ReadLocalSubroutines(data, tableOffset, tableEnd, size, relative, isCff2)
                    : CffIndex.Empty(data);
            }
            if (topDictionary.TryGetInteger(0x0C25, out int fdSelectRelative)) {
                fontDictionaryByGlyph = ReadFdSelect(data, checked(tableOffset + fdSelectRelative), tableEnd, charStrings.Count, fdArray.Count);
            } else if (fdArray.Count > 1) {
                throw new InvalidDataException("A multi-dictionary CFF font is missing FDSelect.");
            }
        } else {
            localSubroutines = new[] { CffIndex.Empty(data) };
        }

        return new OfficeCffFontData(
            isCff2,
            maximumOperandStack,
            charStrings,
            globalSubroutines,
            localSubroutines,
            fontDictionaryByGlyph,
            standardEncodingGlyphs,
            variationStore);
    }

    internal CffSlice GetCharString(int glyphId) {
        if (glyphId < 0 || glyphId >= _charStrings.Count) throw new ArgumentOutOfRangeException(nameof(glyphId));
        return _charStrings[glyphId];
    }

    internal CffIndex GetLocalSubroutines(int glyphId) {
        int dictionary = _fontDictionaryByGlyph == null ? 0 : _fontDictionaryByGlyph[glyphId];
        return _localSubroutines[dictionary];
    }

    internal CffIndex GlobalSubroutines => _globalSubroutines;

    internal int ResolveStandardEncodingGlyph(int characterCode) {
        if (IsCff2 || _standardEncodingGlyphs == null || characterCode < 0 || characterCode >= 256) {
            throw new InvalidDataException("A CFF seac character code is invalid.");
        }
        int glyph = _standardEncodingGlyphs[characterCode];
        if (glyph <= 0 || glyph >= GlyphCount) throw new InvalidDataException("A CFF seac component is absent from the charset.");
        return glyph;
    }

    private static CffIndex ReadLocalSubroutines(
        byte[] data,
        int tableOffset,
        int tableEnd,
        int privateSize,
        int privateRelative,
        bool isCff2) {
        if (privateSize < 0 || privateRelative < 0) throw new InvalidDataException("The CFF Private DICT range is invalid.");
        int privateOffset = checked(tableOffset + privateRelative);
        EnsureRange(privateOffset, privateSize, tableOffset, tableEnd, "The CFF Private DICT is truncated.");
        CffDictionary dictionary = CffDictionary.Parse(data, privateOffset, privateSize);
        if (!dictionary.TryGetInteger(19, out int subrRelative)) return CffIndex.Empty(data);
        int subrOffset = checked(privateOffset + subrRelative);
        EnsureRange(subrOffset, 1, privateOffset, tableEnd, "The CFF local Subr INDEX offset is invalid.");
        return CffIndex.ReadAt(data, subrOffset, tableEnd, isCff2 ? 4 : 2);
    }

    private static int[] ReadFdSelect(byte[] data, int offset, int end, int glyphCount, int fdCount) {
        EnsureRange(offset, 1, 0, end, "The CFF FDSelect table is truncated.");
        int format = data[offset++];
        var result = new int[glyphCount];
        if (format == 0) {
            EnsureRange(offset, glyphCount, 0, end, "The CFF FDSelect format 0 table is truncated.");
            for (int glyph = 0; glyph < glyphCount; glyph++) result[glyph] = ValidateFd(data[offset + glyph], fdCount);
            return result;
        }
        if (format == 3) {
            int rangeCount = ReadUInt16(data, offset, end);
            offset += 2;
            if (rangeCount <= 0) throw new InvalidDataException("The CFF FDSelect range count is invalid.");
            EnsureRange(offset, checked(rangeCount * 3 + 2), 0, end, "The CFF FDSelect format 3 table is truncated.");
            int previousGlyph = -1;
            int previousFd = -1;
            for (int range = 0; range < rangeCount; range++) {
                int firstGlyph = ReadUInt16(data, offset, end);
                int fd = ValidateFd(data[offset + 2], fdCount);
                offset += 3;
                if (range == 0 && firstGlyph != 0 || firstGlyph <= previousGlyph) throw new InvalidDataException("The CFF FDSelect ranges are invalid.");
                if (range > 0) FillFd(result, previousGlyph, firstGlyph, previousFd);
                previousGlyph = firstGlyph;
                previousFd = fd;
            }
            int sentinel = ReadUInt16(data, offset, end);
            if (sentinel != glyphCount) throw new InvalidDataException("The CFF FDSelect sentinel does not match the glyph count.");
            FillFd(result, previousGlyph, sentinel, previousFd);
            return result;
        }
        if (format == 4) {
            uint rangeCountValue = ReadUInt32(data, offset, end);
            offset += 4;
            if (rangeCountValue == 0 || rangeCountValue > int.MaxValue) throw new InvalidDataException("The CFF FDSelect range count is invalid.");
            int rangeCount = (int)rangeCountValue;
            EnsureRange(offset, checked(rangeCount * 6 + 4), 0, end, "The CFF FDSelect format 4 table is truncated.");
            int previousGlyph = -1;
            int previousFd = -1;
            for (int range = 0; range < rangeCount; range++) {
                uint firstValue = ReadUInt32(data, offset, end);
                int fd = ReadUInt16(data, offset + 4, end);
                offset += 6;
                if (firstValue > int.MaxValue || fd >= fdCount) throw new InvalidDataException("The CFF FDSelect range is invalid.");
                int firstGlyph = (int)firstValue;
                if (range == 0 && firstGlyph != 0 || firstGlyph <= previousGlyph) throw new InvalidDataException("The CFF FDSelect ranges are invalid.");
                if (range > 0) FillFd(result, previousGlyph, firstGlyph, previousFd);
                previousGlyph = firstGlyph;
                previousFd = fd;
            }
            uint sentinelValue = ReadUInt32(data, offset, end);
            if (sentinelValue != glyphCount) throw new InvalidDataException("The CFF FDSelect sentinel does not match the glyph count.");
            FillFd(result, previousGlyph, glyphCount, previousFd);
            return result;
        }
        throw new NotSupportedException("The CFF FDSelect format is not supported.");
    }

    private static void FillFd(int[] result, int start, int end, int fd) {
        if (start < 0 || end < start || end > result.Length) throw new InvalidDataException("The CFF FDSelect range is invalid.");
        for (int glyph = start; glyph < end; glyph++) result[glyph] = fd;
    }

    private static int ValidateFd(int value, int count) {
        if (value < 0 || value >= count) throw new InvalidDataException("The CFF FDSelect dictionary index is invalid.");
        return value;
    }

    private static int ReadUInt16(byte[] data, int offset, int end) {
        EnsureRange(offset, 2, 0, end, "CFF data is truncated.");
        return (data[offset] << 8) | data[offset + 1];
    }

    private static uint ReadUInt32(byte[] data, int offset, int end) {
        EnsureRange(offset, 4, 0, end, "CFF data is truncated.");
        return ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];
    }

    private static void EnsureRange(int offset, int length, int lowerBound, int upperBound, string message) {
        if (offset < lowerBound || length < 0 || offset > upperBound - length) throw new InvalidDataException(message);
    }

    internal readonly struct CffSlice {
        internal CffSlice(byte[] data, int offset, int length) {
            Data = data;
            Offset = offset;
            Length = length;
        }

        internal byte[] Data { get; }
        internal int Offset { get; }
        internal int Length { get; }
    }

    internal sealed class CffIndex {
        private readonly byte[] _data;
        private readonly int[] _offsets;
        private readonly int _dataOffset;

        private CffIndex(byte[] data, int[] offsets, int dataOffset) {
            _data = data;
            _offsets = offsets;
            _dataOffset = dataOffset;
        }

        internal int Count => _offsets.Length == 0 ? 0 : _offsets.Length - 1;

        internal CffSlice this[int index] {
            get {
                if (index < 0 || index >= Count) throw new ArgumentOutOfRangeException(nameof(index));
                int start = checked(_dataOffset + _offsets[index] - 1);
                int length = _offsets[index + 1] - _offsets[index];
                return new CffSlice(_data, start, length);
            }
        }

        internal static CffIndex Empty(byte[] data) => new(data, Array.Empty<int>(), 0);

        internal static CffIndex ReadAt(byte[] data, int offset, int end, int countSize) {
            int cursor = offset;
            return Read(data, ref cursor, end, countSize);
        }

        internal static CffIndex Read(byte[] data, ref int cursor, int end, int countSize) {
            if (countSize != 2 && countSize != 4) throw new ArgumentOutOfRangeException(nameof(countSize));
            EnsureRange(cursor, countSize, 0, end, "The CFF INDEX count is truncated.");
            uint countValue = countSize == 2
                ? (uint)ReadUInt16(data, cursor, end)
                : ReadUInt32(data, cursor, end);
            cursor += countSize;
            if (countValue == 0) return Empty(data);
            if (countValue > MaximumIndexObjects || countValue > int.MaxValue) throw new InvalidDataException("The CFF INDEX object count is invalid.");
            int count = (int)countValue;
            EnsureRange(cursor, 1, 0, end, "The CFF INDEX offSize is truncated.");
            int offSize = data[cursor++];
            if (offSize < 1 || offSize > 4) throw new InvalidDataException("The CFF INDEX offSize is invalid.");
            EnsureRange(cursor, checked((count + 1) * offSize), 0, end, "The CFF INDEX offsets are truncated.");
            var offsets = new int[count + 1];
            for (int index = 0; index <= count; index++) {
                uint value = 0;
                for (int part = 0; part < offSize; part++) value = (value << 8) | data[cursor++];
                if (value == 0 || value > int.MaxValue || index > 0 && value < offsets[index - 1]) {
                    throw new InvalidDataException("The CFF INDEX offsets are invalid.");
                }
                offsets[index] = (int)value;
            }
            if (offsets[0] != 1) throw new InvalidDataException("The first CFF INDEX offset must be one.");
            int dataLength = offsets[count] - 1;
            EnsureRange(cursor, dataLength, 0, end, "The CFF INDEX object data is truncated.");
            int dataOffset = cursor;
            cursor += dataLength;
            return new CffIndex(data, offsets, dataOffset);
        }
    }

    private sealed class CffDictionary {
        private readonly Dictionary<int, double[]> _values;

        private CffDictionary(Dictionary<int, double[]> values) => _values = values;

        internal static CffDictionary Parse(byte[] data, int offset, int length) {
            EnsureRange(offset, length, 0, data.Length, "The CFF DICT data is truncated.");
            int end = checked(offset + length);
            var values = new Dictionary<int, double[]>();
            var operands = new List<double>();
            while (offset < end) {
                int value = data[offset];
                if (value <= 27) {
                    offset++;
                    int operation = value == 12
                        ? 0x0C00 | ReadOperatorByte(data, ref offset, end)
                        : value;
                    values[operation] = operands.ToArray();
                    operands.Clear();
                } else {
                    operands.Add(ReadNumber(data, ref offset, end, charString: false));
                }
            }
            if (operands.Count != 0) throw new InvalidDataException("The CFF DICT ends with unconsumed operands.");
            return new CffDictionary(values);
        }

        internal int GetRequiredInteger(int operation, string name) {
            if (!TryGetInteger(operation, out int value)) throw new InvalidDataException("The CFF Top DICT is missing " + name + ".");
            return value;
        }

        internal bool ContainsOperation(int operation) => _values.ContainsKey(operation);

        internal bool HasNonNegativeIntegerOperands(int operation, int count) {
            if (!_values.TryGetValue(operation, out double[]? operands) || operands.Length != count) return false;
            for (int index = 0; index < operands.Length; index++) {
                double operand = operands[index];
                if (operand < 0D || operand > int.MaxValue || operand != Math.Truncate(operand)) return false;
            }
            return true;
        }

        internal bool TryGetInteger(int operation, out int value) {
            value = 0;
            if (!_values.TryGetValue(operation, out double[]? operands) || operands.Length < 1) return false;
            double candidate = operands[operands.Length - 1];
            if (candidate < 0 || candidate > int.MaxValue || candidate != Math.Truncate(candidate)) {
                throw new InvalidDataException("A CFF DICT offset is not a non-negative integer.");
            }
            value = checked((int)candidate);
            return true;
        }

        internal bool TryGetPair(int operation, out int first, out int second) {
            first = second = 0;
            if (!_values.TryGetValue(operation, out double[]? operands) || operands.Length < 2) return false;
            double firstValue = operands[operands.Length - 2];
            double secondValue = operands[operands.Length - 1];
            if (firstValue < 0 || secondValue < 0 || firstValue > int.MaxValue || secondValue > int.MaxValue
                || firstValue != Math.Truncate(firstValue) || secondValue != Math.Truncate(secondValue)) {
                throw new InvalidDataException("A CFF DICT range is invalid.");
            }
            first = checked((int)firstValue);
            second = checked((int)secondValue);
            return true;
        }
    }

    internal static double ReadNumber(byte[] data, ref int offset, int end, bool charString) {
        EnsureRange(offset, 1, 0, end, "A CFF number is truncated.");
        int first = data[offset++];
        if (first >= 32 && first <= 246) return first - 139;
        if (first >= 247 && first <= 250) {
            EnsureRange(offset, 1, 0, end, "A CFF number is truncated.");
            return (first - 247) * 256 + data[offset++] + 108;
        }
        if (first >= 251 && first <= 254) {
            EnsureRange(offset, 1, 0, end, "A CFF number is truncated.");
            return -((first - 251) * 256) - data[offset++] - 108;
        }
        if (first == 28) {
            int value = ReadUInt16(data, offset, end);
            offset += 2;
            return unchecked((short)value);
        }
        if (first == 29 && !charString) {
            uint value = ReadUInt32(data, offset, end);
            offset += 4;
            return unchecked((int)value);
        }
        if (first == 30 && !charString) return ReadReal(data, ref offset, end);
        if (first == 255) {
            uint value = ReadUInt32(data, offset, end);
            offset += 4;
            return unchecked((int)value) / 65536D;
        }
        throw new InvalidDataException("The CFF number encoding is invalid.");
    }

    private static double ReadReal(byte[] data, ref int offset, int end) {
        var text = new StringBuilder();
        bool done = false;
        while (!done) {
            EnsureRange(offset, 1, 0, end, "A CFF real number is truncated.");
            int value = data[offset++];
            done = AppendNibble(text, value >> 4) || AppendNibble(text, value & 0x0F);
        }
        if (!double.TryParse(text.ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out double result)
            || double.IsNaN(result) || double.IsInfinity(result)) {
            throw new InvalidDataException("A CFF real number is invalid.");
        }
        return result;
    }

    private static bool AppendNibble(StringBuilder text, int value) {
        if (value <= 9) text.Append((char)('0' + value));
        else if (value == 10) text.Append('.');
        else if (value == 11) text.Append('E');
        else if (value == 12) text.Append("E-");
        else if (value == 14) text.Append('-');
        else if (value == 15) return true;
        else throw new InvalidDataException("A CFF real-number nibble is reserved.");
        return false;
    }

    private static int ReadOperatorByte(byte[] data, ref int offset, int end) {
        EnsureRange(offset, 1, 0, end, "A CFF escaped operator is truncated.");
        return data[offset++];
    }
}

/// <summary>Precomputed CFF2 variation-region scalars for each ItemVariationData selection.</summary>
internal sealed class OfficeCffVariationStore {
    private readonly double[]?[] _scalars;

    private OfficeCffVariationStore(double[]?[] scalars) => _scalars = scalars;

    internal static OfficeCffVariationStore Parse(
        OfficeOpenTypeReader reader,
        int offset,
        int cffEnd,
        OfficeFontVariationModel variations) {
        if (offset < 0 || offset > cffEnd - 2) throw new InvalidDataException("The CFF2 VariationStore is truncated.");
        int length = reader.ReadUInt16(offset);
        int itemStore = checked(offset + 2);
        if (length < 8 || itemStore > cffEnd - length) throw new InvalidDataException("The CFF2 VariationStore length is invalid.");
        int storeEnd = checked(itemStore + length);
        if (reader.ReadUInt16(itemStore) != 1) throw new NotSupportedException("The CFF2 ItemVariationStore format is not supported.");
        uint regionListRelative = reader.ReadUInt32(itemStore + 2);
        int dataCount = reader.ReadUInt16(itemStore + 6);
        if (dataCount <= 0 || dataCount > 4096 || itemStore + 8 > storeEnd - dataCount * 4) {
            throw new InvalidDataException("The CFF2 ItemVariationStore directory is invalid.");
        }
        int regionList = checked(itemStore + (int)regionListRelative);
        if (regionList < itemStore || regionList > storeEnd - 4) throw new InvalidDataException("The CFF2 VariationRegionList offset is invalid.");
        int axisCount = reader.ReadUInt16(regionList);
        int regionCount = reader.ReadUInt16(regionList + 2);
        if (axisCount != variations.AxisCount || regionCount <= 0 || regionCount > 32768
            || regionList + 4 > storeEnd - checked(axisCount * regionCount * 6)) {
            throw new InvalidDataException("The CFF2 VariationRegionList dimensions are invalid.");
        }
        var regionScalars = new double[regionCount];
        int regionCursor = regionList + 4;
        for (int region = 0; region < regionCount; region++) {
            double scalar = 1D;
            for (int axis = 0; axis < axisCount; axis++) {
                double start = reader.ReadF2Dot14(regionCursor);
                double peak = reader.ReadF2Dot14(regionCursor + 2);
                double end = reader.ReadF2Dot14(regionCursor + 4);
                regionCursor += 6;
                scalar *= OfficeOpenTypeVariationRegion.CalculateScalar(
                    variations.NormalizedCoordinates[axis],
                    start,
                    peak,
                    end);
            }
            regionScalars[region] = scalar;
        }

        var selections = new double[]?[dataCount];
        var selectionsByOffset = new Dictionary<int, double[]>();
        int persistentScalarCount = 0;
        int maximumPersistentScalars = Math.Min(
            1_000_000,
            Math.Max(64, reader.Data.Length / 2));
        for (int selection = 0; selection < dataCount; selection++) {
            uint relative = reader.ReadUInt32(itemStore + 8 + selection * 4);
            if (relative == 0) {
                selections[selection] = null;
                continue;
            }
            if (relative > int.MaxValue) throw new InvalidDataException("A CFF2 ItemVariationData offset is invalid.");
            int dataOffset = checked(itemStore + (int)relative);
            if (dataOffset < itemStore || dataOffset > storeEnd - 6) throw new InvalidDataException("A CFF2 ItemVariationData offset is invalid.");
            if (selectionsByOffset.TryGetValue(dataOffset, out double[]? existingSelection)) {
                selections[selection] = existingSelection;
                continue;
            }
            _ = reader.ReadUInt16(dataOffset); // itemCount is not consumed by CFF2 CharString blending.
            _ = reader.ReadUInt16(dataOffset + 2); // shortDeltaCount is not consumed here.
            int indexCount = reader.ReadUInt16(dataOffset + 4);
            if (indexCount < 0 || dataOffset + 6 > storeEnd - indexCount * 2) throw new InvalidDataException("A CFF2 region-index array is truncated.");
            if (indexCount > maximumPersistentScalars - persistentScalarCount) {
                throw new InvalidDataException("CFF2 variation metadata exceeds the bounded allocation budget.");
            }
            persistentScalarCount += indexCount;
            var selected = new double[indexCount];
            for (int index = 0; index < indexCount; index++) {
                int regionIndex = reader.ReadUInt16(dataOffset + 6 + index * 2);
                if (regionIndex >= regionScalars.Length) throw new InvalidDataException("A CFF2 region index is outside the VariationRegionList.");
                selected[index] = regionScalars[regionIndex];
            }
            selectionsByOffset.Add(dataOffset, selected);
            selections[selection] = selected;
        }
        return new OfficeCffVariationStore(selections);
    }

    internal IReadOnlyList<double> GetScalars(int variationDataIndex) {
        if (variationDataIndex < 0 || variationDataIndex >= _scalars.Length) throw new InvalidDataException("The CFF2 vsindex value is invalid.");
        return _scalars[variationDataIndex]
            ?? throw new InvalidDataException("The CFF2 vsindex value references a null ItemVariationData entry.");
    }

}
