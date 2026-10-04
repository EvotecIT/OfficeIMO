using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;

namespace OfficeIMO.Core.Internal;

/// <summary>Reassembles bounded OPC interleaved physical items into logical parts after ZIP validation.</summary>
internal static class OfficeOpcPieceAssembler {
    internal static bool HasPieceSuffix(string name) {
        int slash = name.LastIndexOf('/');
        return slash >= 0 && name.Substring(slash + 1).StartsWith("[", StringComparison.Ordinal) && name.EndsWith(".piece", StringComparison.OrdinalIgnoreCase);
    }
    internal static Dictionary<string, byte[]> Assemble(Dictionary<string, byte[]> items, int maximumPartBytes, CancellationToken token) {
        var parts = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        var sets = new Dictionary<string, SortedDictionary<int, Piece>>(StringComparer.OrdinalIgnoreCase);
        foreach (var item in items) {
            token.ThrowIfCancellationRequested();
            int slash = item.Key.LastIndexOf('/');
            string suffix = item.Key.Substring(slash + 1);
            if (!HasPieceSuffix(item.Key)) {
                parts.Add(item.Key, item.Value); continue;
            }
            int close = suffix.IndexOf(']');
            bool last = close >= 0 && suffix.Substring(close + 1).Equals(".last.piece", StringComparison.OrdinalIgnoreCase);
            if (close < 2 || (!last && !suffix.Substring(close + 1).Equals(".piece", StringComparison.OrdinalIgnoreCase)) ||
                !int.TryParse(suffix.Substring(1, close - 1), NumberStyles.None, CultureInfo.InvariantCulture, out int index))
                throw new InvalidDataException("Invalid OPC piece name.");
            string name = item.Key.Substring(0, slash);
            if (HasPieceSuffix(name)) throw new InvalidDataException("Nested OPC piece storage is invalid.");
            if (!sets.TryGetValue(name, out var pieces)) { pieces = new SortedDictionary<int, Piece>(); sets.Add(name, pieces); }
            if (pieces.ContainsKey(index)) throw new InvalidDataException("Duplicate OPC piece index.");
            pieces.Add(index, new Piece(item.Value, last));
        }
        foreach (var set in sets) {
            token.ThrowIfCancellationRequested();
            if (parts.ContainsKey(set.Key)) throw new InvalidDataException("OPC part has both atomic and interleaved storage.");
            long total = 0; int expected = 0;
            foreach (var piece in set.Value) {
                if (piece.Key != expected || piece.Value.Last != (expected == set.Value.Count - 1))
                    throw new InvalidDataException("OPC pieces require consecutive indices and one terminal piece.");
                total += piece.Value.Bytes.Length;
                if (total > maximumPartBytes) throw new InvalidDataException("Assembled OPC part exceeds the part byte limit.");
                expected++;
            }
            var data = new byte[(int)total]; int position = 0;
            foreach (var piece in set.Value.Values) {
                token.ThrowIfCancellationRequested();
                Buffer.BlockCopy(piece.Bytes, 0, data, position, piece.Bytes.Length); position += piece.Bytes.Length;
            }
            parts.Add(set.Key, data);
        }
        return parts;
    }
    private sealed class Piece {
        internal readonly byte[] Bytes;
        internal readonly bool Last;
        internal Piece(byte[] bytes, bool last) { Bytes = bytes; Last = last; }
    }
}
