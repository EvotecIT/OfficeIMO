using System.Threading;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Visio {
    /// <summary>Reads the native pointer graph once; container, expansion and traversal budgets are independent.</summary>
    internal sealed class VisioBinaryContainer {
        private readonly byte[] _source;
        private readonly VisioLegacyBinaryImportOptions _options;
        private readonly CancellationToken _token;
        private readonly Dictionary<uint, Node> _nodes = new();
        private readonly HashSet<uint> _active = new();
        private int _expandedBytes, _records;

        internal VisioBinaryContainer(byte[] bytes, VisioLegacyBinaryImportOptions options, CancellationToken token) {
            _options = options; _token = token;
            var limits = options.Limits;
            var compoundOptions = new OfficeCompoundReadOptions(
                maxDirectoryEntries: (int)Math.Min(int.MaxValue, (long)limits.MaxCompoundStreams * 4 + 1),
                maxStreamCount: limits.MaxCompoundStreams,
                maxStreamBytes: limits.MaxInputBytes, maxTotalStreamBytes: limits.MaxInputBytes);
            if (!OfficeCompoundFileReader.TryRead(bytes, compoundOptions, token, out OfficeCompoundFile? compound, out string? error))
                throw new InvalidDataException("Invalid legacy Visio compound file: " + error);
            if (!compound!.Streams.TryGetValue("VisioDocument", out _source!))
                throw new InvalidDataException("The compound file has no root VisioDocument stream.");
            DirectoryEntries = compound.Entries;
            byte[] magic = System.Text.Encoding.ASCII.GetBytes("Visio (TM) Drawing\r\n\0");
            if (_source.Length < 54 || !magic.SequenceEqual(_source.Take(magic.Length)))
                throw new InvalidDataException("Invalid VisioDocument stream signature or header.");
            Version = _source[26];
            if (Version != 11)
                throw new NotSupportedException($"Binary Visio generation {Version} is outside the supported version 11 profile.");
        }

        internal int Version { get; }
        internal IReadOnlyList<OfficeCompoundFileEntry> DirectoryEntries { get; }
        internal Node ReadRoot() => ReadPointer(_source, 36, 0, 0);
        internal void CheckCancellation() => _token.ThrowIfCancellationRequested();

        internal void CountRecord() => CountRecords(1);

        private void CountRecords(int count) {
            _token.ThrowIfCancellationRequested();
            if (count > _options.Limits.MaxRecords - _records) throw new InvalidDataException("Binary Visio record/pointer budget exceeded.");
            _records += count;
        }

        private Node ReadPointer(byte[] data, int offset, uint id, int depth) {
            CountRecord();
            if (depth > _options.MaxDepth) throw new InvalidDataException("Binary Visio pointer depth exceeded.");
            VisioBinaryData.Require(data, offset, 18);
            uint type = VisioBinaryData.U32(data, offset);
            uint position = VisioBinaryData.U32(data, offset + 8);
            int length = VisioBinaryData.Size(data, offset + 12);
            ushort format = VisioBinaryData.U16(data, offset + 16);
            int start = VisioBinaryData.Size(position);
            VisioBinaryData.Require(_source, start, length);
            // Empty native slots can share offset zero with different metadata.
            if (length == 0) return new Node(type, id, format, 0, Array.Empty<byte>(), 0);
            if (!_active.Add(position)) throw new InvalidDataException("Binary Visio pointer graph contains a cycle.");
            try {
                if (_nodes.TryGetValue(position, out Node? cached)) {
                    if (cached.Type != type || cached.Format != format || cached.SourceLength != length)
                        throw new InvalidDataException("Binary Visio pointers disagree about a shared stream.");
                    if ((long)depth + cached.DescendantDepth > _options.MaxDepth)
                        throw new InvalidDataException("Binary Visio pointer depth exceeded.");
                    // Aliases share storage, but their logical traversal still consumes the budget.
                    CountRecords(cached.DescendantRecords);
                    return cached.WithId(id);
                }
                bool compressed = (format & 2) != 0;
                byte[] decoded = compressed ? Expand(start, length) : Copy(start, length);
                var node = new Node(type, id, format, length, decoded, compressed ? 4 : 0);
                _nodes.Add(position, node);
                if ((format >> 4) == 5 && type != 0x16) ReadChildren(node, depth);
                return node;
            } finally { _active.Remove(position); }
        }

        private void ReadChildren(Node node, int depth) {
            int recordsBefore = _records;
            byte[] data = node.Data;
            long tableValue = (long)VisioBinaryData.U32(data, node.Shift) + node.Shift - 4;
            if (tableValue < 0 || tableValue > data.Length) throw new InvalidDataException("Binary Visio pointer table offset is invalid.");
            int table = (int)tableValue;
            VisioBinaryData.Require(data, table, 12);
            int orderCount = VisioBinaryData.Size(data, table);
            int count = VisioBinaryData.Size(data, table + 4);
            if (count > _options.Limits.MaxRecords - _records)
                throw new InvalidDataException("Binary Visio pointer budget exceeded.");
            int first = table + 12;
            if ((long)count * 18 > data.Length - first) throw new InvalidDataException("Binary Visio pointer table is truncated.");
            var children = new Dictionary<uint, Node>();
            for (int index = 0; index < count; index++) {
                int pointer = first + index * 18;
                if (VisioBinaryData.U32(data, pointer) != 0)
                    children.Add((uint)index, ReadPointer(data, pointer, (uint)index, depth + 1));
                else CountRecord();
            }
            int orderStart = first + count * 18;
            if (orderCount <= 1) orderCount = 0;
            if (orderCount > _options.Limits.MaxRecords - _records) throw new InvalidDataException("Binary Visio pointer order budget exceeded.");
            if ((long)orderCount * 4 > data.Length - orderStart)
                throw new InvalidDataException("Binary Visio pointer order is truncated.");
            var emitted = new HashSet<uint>();
            var ordered = new List<Node>(children.Count);
            for (int index = 0; index < orderCount; index++) {
                CountRecord();
                uint id = VisioBinaryData.U32(data, orderStart + index * 4);
                if (children.TryGetValue(id, out Node? child) && emitted.Add(id)) ordered.Add(child);
            }
            foreach (var child in children.OrderBy(pair => pair.Key)) if (emitted.Add(child.Key)) ordered.Add(child.Value);
            node.CompleteChildren(ordered.ToArray(), _records - recordsBefore);
        }

        private byte[] Copy(int start, int length) {
            Charge(length);
            var data = new byte[length];
            Buffer.BlockCopy(_source, start, data, 0, length);
            return data;
        }

        private byte[] Expand(int start, int length) {
            int end = checked(start + length);
            var history = new byte[4096];
            using var output = new MemoryStream();
            while (start < end) {
                _token.ThrowIfCancellationRequested();
                byte flags = _source[start++];
                if (start == end) break; // Native streams can end with a flag byte and no items.
                for (int bit = 0; bit < 8 && start < end; bit++) {
                    if ((flags & (1 << bit)) != 0) {
                        Append(_source[start++]);
                    } else {
                        if (end - start < 2) throw new InvalidDataException("Binary Visio compressed match is truncated.");
                        int low = _source[start++], high = _source[start++];
                        int address = ((low | ((high & 0xf0) << 4)) + 18) & 4095;
                        int count = (high & 15) + 3;
                        for (int item = 0; item < count; item++) Append(history[(address + item) & 4095]);
                    }
                }
            }
            return output.ToArray();

            void Append(byte value) {
                Charge(1);
                history[(int)output.Length & 4095] = value;
                output.WriteByte(value);
            }
        }

        private void Charge(int count) {
            if (count > _options.MaxDecompressedBytes - _expandedBytes)
                throw new InvalidDataException("Binary Visio cumulative decoded-byte budget exceeded.");
            _expandedBytes += count;
        }

        internal sealed class Node {
            internal Node(uint type, uint id, ushort format, int length, byte[] data, int shift) {
                Type = type; Id = id; Format = format; SourceLength = length; Data = data; Shift = shift;
            }
            internal uint Type { get; }
            internal uint Id { get; }
            internal ushort Format { get; }
            internal int SourceLength { get; }
            internal byte[] Data { get; }
            internal int Shift { get; }
            internal IReadOnlyList<Node> Children { get; private set; } = Array.Empty<Node>();
            internal int DescendantRecords { get; private set; }
            internal int DescendantDepth { get; private set; }
            internal void CompleteChildren(Node[] children, int records) {
                Children = children;
                DescendantRecords = records;
                DescendantDepth = children.Length == 0 ? 0 : children.Max(child => child.DescendantDepth) + 1;
            }
            internal Node WithId(uint id) {
                if (Id == id) return this;
                // Only completed nodes can be aliased: the active-position guard rejects cycles.
                return new Node(Type, id, Format, SourceLength, Data, Shift) {
                    Children = Children, DescendantRecords = DescendantRecords, DescendantDepth = DescendantDepth
                };
            }
        }
    }

    internal static class VisioBinaryData {
        internal static void Require(byte[] bytes, int offset, int length) {
            if (offset < 0 || length < 0 || offset > bytes.Length - length)
                throw new InvalidDataException("Binary Visio data extends beyond its containing record.");
        }
        internal static ushort U16(byte[] bytes, int offset) {
            Require(bytes, offset, 2);
            return (ushort)(bytes[offset] | bytes[offset + 1] << 8);
        }
        internal static uint U32(byte[] bytes, int offset) {
            Require(bytes, offset, 4);
            return (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24);
        }
        internal static int Size(byte[] bytes, int offset) => Size(U32(bytes, offset));
        internal static int Size(uint value) => value <= int.MaxValue ? (int)value
            : throw new InvalidDataException("Binary Visio length exceeds the managed input range.");
        internal static double Number(byte[] bytes, int offset) {
            Require(bytes, offset, 8);
            var bits = (ulong)U32(bytes, offset) | ((ulong)U32(bytes, offset + 4) << 32);
            double value = BitConverter.Int64BitsToDouble(unchecked((long)bits));
            if (double.IsNaN(value) || double.IsInfinity(value)) throw new InvalidDataException("Binary Visio contains a nonfinite cached number.");
            return value;
        }
    }
}
