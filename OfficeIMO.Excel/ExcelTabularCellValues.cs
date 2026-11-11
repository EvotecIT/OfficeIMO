namespace OfficeIMO.Excel {
    /// <summary>
    /// Row-major cell values captured once from caller selectors and kept until a deferred save.
    /// Text, double, int, long, DateTime and bool are stored without boxing; any other value keeps its
    /// object, so <see cref="GetValue"/> returns exactly what the selector returned. Storage grows in
    /// chunks small enough to stay off the large object heap.
    /// </summary>
    internal sealed class ExcelTabularCellValues {
        private const int ChunkBits = 13;
        private const int ChunkSize = 1 << ChunkBits;
        private const int ChunkMask = ChunkSize - 1;
        private const long TicksMask = 0x3FFF_FFFF_FFFF_FFFFL;
        private const int DateKindShift = 62;

        private readonly byte[][] _kinds;
        private readonly long[][] _bits;
        private object?[][] _references = new object?[4][];
        private int _referenceCount;

        internal ExcelTabularCellValues(int columnCount, int rowCount) {
            ColumnCount = columnCount;
            RowCount = rowCount;
            int cellCount = checked(columnCount * rowCount);
            int chunkCount = (cellCount + ChunkMask) >> ChunkBits;
            _kinds = new byte[chunkCount][];
            _bits = new long[chunkCount][];
            for (int chunk = 0; chunk < chunkCount; chunk++) {
                int length = Math.Min(ChunkSize, cellCount - (chunk << ChunkBits));
                _kinds[chunk] = new byte[length];
                _bits[chunk] = new long[length];
            }
        }

        internal int ColumnCount { get; }

        internal int RowCount { get; }

        /// <summary>Stores a value; text and other references are appended in call order.</summary>
        internal void Set(int index, object? value) {
            ExcelBufferedCellKind kind;
            long bits;
            switch (value) {
                case null:
                    kind = ExcelBufferedCellKind.Blank;
                    bits = 0L;
                    break;
                case string text:
                    kind = ExcelBufferedCellKind.Text;
                    bits = AddReference(text);
                    break;
                case double number:
                    kind = ExcelBufferedCellKind.Number;
                    bits = BitConverter.DoubleToInt64Bits(number);
                    break;
                case int number:
                    kind = ExcelBufferedCellKind.Int32;
                    bits = number;
                    break;
                case long number:
                    kind = ExcelBufferedCellKind.Int64;
                    bits = number;
                    break;
                case DateTime date:
                    kind = ExcelBufferedCellKind.DateTime;
                    bits = date.Ticks | ((long)date.Kind << DateKindShift);
                    break;
                case bool boolean:
                    kind = ExcelBufferedCellKind.Boolean;
                    bits = boolean ? 1L : 0L;
                    break;
                default:
                    kind = ExcelBufferedCellKind.Value;
                    bits = AddReference(value);
                    break;
            }

            _kinds[index >> ChunkBits][index & ChunkMask] = (byte)kind;
            _bits[index >> ChunkBits][index & ChunkMask] = bits;
        }

        private long AddReference(object value) {
            int chunk = _referenceCount >> ChunkBits;
            if (chunk == _references.Length) Array.Resize(ref _references, _references.Length * 2);
            object?[] references = _references[chunk] ??= new object?[ChunkSize];
            references[_referenceCount & ChunkMask] = value;
            return _referenceCount++;
        }

        /// <summary>Exposes the chunk holding <paramref name="index"/> and the cell's offset in it, for sequential readers.</summary>
        internal int GetChunk(int index, out byte[] kinds, out long[] bits) {
            kinds = _kinds[index >> ChunkBits];
            bits = _bits[index >> ChunkBits];
            return index & ChunkMask;
        }

        /// <summary>Resolves the reference id stored in a Text or Value cell's bits.</summary>
        internal object GetReferenceById(long reference) =>
            _references[(int)(reference >> ChunkBits)][(int)reference & ChunkMask]!;

        internal ExcelBufferedCellKind GetKind(int index) => (ExcelBufferedCellKind)_kinds[index >> ChunkBits][index & ChunkMask];

        internal long GetBits(int index) => _bits[index >> ChunkBits][index & ChunkMask];

        /// <summary>Gets the string or object of a Text or Value cell.</summary>
        internal object GetReference(int index) {
            long reference = GetBits(index);
            return _references[(int)(reference >> ChunkBits)][(int)reference & ChunkMask]!;
        }

        internal double GetDouble(int index) => BitConverter.Int64BitsToDouble(GetBits(index));

        internal DateTime GetDateTime(int index) {
            long bits = GetBits(index);
            return new DateTime(bits & TicksMask, (DateTimeKind)((ulong)bits >> DateKindShift));
        }

        /// <summary>Returns the value the selector produced, boxed again for value types.</summary>
        internal object? GetValue(int index) {
            switch (GetKind(index)) {
                case ExcelBufferedCellKind.Text:
                case ExcelBufferedCellKind.Value:
                    return GetReference(index);
                case ExcelBufferedCellKind.Number:
                    return GetDouble(index);
                case ExcelBufferedCellKind.Int32:
                    return (int)GetBits(index);
                case ExcelBufferedCellKind.Int64:
                    return GetBits(index);
                case ExcelBufferedCellKind.DateTime:
                    return GetDateTime(index);
                case ExcelBufferedCellKind.Boolean:
                    return GetBits(index) != 0L;
                default:
                    return null;
            }
        }

        /// <summary>Matches <see cref="ExcelDirectTabularValue.Normalize(object?, bool)"/> for the stored value.</summary>
        internal ExcelDirectTabularValue Normalize(int index, bool preserveMissingValues) {
            switch (GetKind(index)) {
                case ExcelBufferedCellKind.Blank:
                    return ExcelDirectTabularValue.Normalize(null, preserveMissingValues);
                case ExcelBufferedCellKind.Text:
                    return ExcelDirectTabularValue.FromText((string)GetReference(index));
                case ExcelBufferedCellKind.Number: {
                    double number = GetDouble(index);
                    return double.IsNaN(number) || double.IsInfinity(number)
                        ? ExcelDirectTabularValue.Unsupported
                        : ExcelDirectTabularValue.FromNumber(number);
                }
                case ExcelBufferedCellKind.Int32:
                    return ExcelDirectTabularValue.FromNumber((int)GetBits(index));
                case ExcelBufferedCellKind.Int64:
                    return ExcelDirectTabularValue.FromNumber(GetBits(index));
                case ExcelBufferedCellKind.Boolean:
                    return ExcelDirectTabularValue.FromBoolean(GetBits(index) != 0L);
                case ExcelBufferedCellKind.DateTime:
                    return ExcelDirectTabularValue.Unsupported;
                default:
                    return ExcelDirectTabularValue.Normalize(GetReference(index), preserveMissingValues);
            }
        }

        /// <summary>Matches <see cref="ExcelDirectTabularValue.Normalize(object?, ExcelDateSystem, bool)"/> for the stored value.</summary>
        internal ExcelDirectTabularValue Normalize(int index, ExcelDateSystem dateSystem, bool preserveMissingValues) =>
            GetKind(index) == ExcelBufferedCellKind.DateTime
                ? ExcelDirectTabularValue.FromDate(ExcelDateSystemConverter.ToSerial(GetDateTime(index), dateSystem))
                : Normalize(index, preserveMissingValues);
    }
}
