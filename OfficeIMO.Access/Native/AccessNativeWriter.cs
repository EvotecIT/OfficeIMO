using System.Text;

namespace OfficeIMO.Access;

/// <summary>Seed-free Jet4/ACE12 page construction. The complete bounded plan is validated before destination I/O.</summary>
internal sealed partial class AccessNativeWriter {
    private const int PageSize = 4096;
    private readonly List<byte[]> _pages = new List<byte[]>();
    private readonly List<Table> _tables = new List<Table>();
    private readonly long _maximumBytes;
    private readonly CancellationToken _cancellation;
    private sealed class Column {
        internal Column(string name, byte type, int size, bool variable = false) { Name = name; Type = type; Size = size; Variable = variable; }
        internal readonly string Name;
        internal readonly byte Type;
        internal readonly int Size;
        internal readonly bool Variable;
        internal bool AutoNumber;
        internal bool IsSystemSid;
        internal byte Precision, Scale;
        internal readonly List<int> LongPages = new List<int>();
    }
    private sealed class Index {
        internal Index(string name, int[] columns, byte flags, byte type = 0) { Name = name; Columns = columns; Flags = flags; Type = type; }
        internal readonly string Name;
        internal readonly int[] Columns;
        internal readonly byte Flags, Type;
        internal int RootPage, RelatedTable, RelatedIndex = -1, UniqueCount;
        internal byte RelatedType;
        internal readonly List<int> Pages = new List<int>();
    }
    private sealed class Table {
        internal Table(int definition, string name, Column[] columns, object?[][] rows) { DefinitionPage = definition; Name = name; Columns = columns; Rows = rows; }
        internal readonly int DefinitionPage;
        internal readonly string Name;
        internal readonly Column[] Columns;
        internal readonly object?[][] Rows;
        internal int MapPage, AutoNumberLast;
        internal readonly List<Index> Indexes = new List<Index>();
        internal readonly List<int> DataPages = new List<int>();
        internal readonly List<uint> RowIds = new List<uint>();
        internal readonly List<uint> MapPointers = new List<uint>();
    }
    private AccessNativeWriter(long maximumBytes, CancellationToken cancellation) { _maximumBytes = maximumBytes; _cancellation = cancellation; }
    internal static AccessNativeWriter Build(AccessDocument document, long maximumBytes, CancellationToken cancellation) {
        var writer = new AccessNativeWriter(maximumBytes, cancellation);
        writer.Initialize(document); return writer;
    }
    internal long Length => (long)_pages.Count * PageSize;
    internal void Write(Stream destination, CancellationToken cancellation) {
        foreach (byte[] page in _pages) { cancellation.ThrowIfCancellationRequested(); destination.Write(page, 0, page.Length); }
    }
    private int Allocate() {
        _cancellation.ThrowIfCancellationRequested();
        if ((long)(_pages.Count + 1) * PageSize > _maximumBytes) throw new InvalidDataException("Native Access output exceeds MaxOutputBytes.");
        _pages.Add(new byte[PageSize]); return _pages.Count - 1;
    }
    private static Index[] PhysicalIndexes(Table table) => table.Indexes.GroupBy(x => x.RootPage).Select(x => x.First()).ToArray();
    private static void U16(byte[] bytes, int offset, ushort value) { bytes[offset] = (byte)value; bytes[offset + 1] = (byte)(value >> 8); }
    private static void U32(byte[] bytes, int offset, uint value) { for (int i = 0; i < 4; i++) bytes[offset + i] = (byte)(value >> (8 * i)); }
    private static void BigEndian(byte[] bytes, int offset, uint value, int count = 4) { for (int i = 0; i < count; i++) bytes[offset + i] = (byte)(value >> ((count - i - 1) * 8)); }
}
