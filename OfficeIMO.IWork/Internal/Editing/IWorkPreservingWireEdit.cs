using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Replaces one existing fixed-width scalar without reserializing unknown protobuf fields.</summary>
internal static class IWorkPreservingWireEdit {
    internal static void ReplaceFloat(byte[] payload, IReadOnlyList<int> path, float value,
        IWorkReadOptions limits, CancellationToken token) {
        var span = FindSpan(payload, path.Select(field => (field, -1)).ToArray(), 5, limits, token);
        byte[] bytes = BitConverter.GetBytes(value); if (!BitConverter.IsLittleEndian) Array.Reverse(bytes);
        Buffer.BlockCopy(bytes, 0, payload, span.Offset, 4);
    }

    /// <summary>Patches a fixed-width range inside an existing bytes field. Occurrence -1 requires a unique field.</summary>
    internal static void ReplaceBytesRange(byte[] payload, IReadOnlyList<(int Field, int Occurrence)> path,
        int valueOffset, byte[] value, IWorkReadOptions limits, CancellationToken token) {
        var span = FindSpan(payload, path, 2, limits, token);
        if (valueOffset < 0 || valueOffset > span.Length || value.Length > span.Length - valueOffset)
            throw new NotSupportedException("The scalar edit exceeds its existing bytes field.");
        token.ThrowIfCancellationRequested();
        Buffer.BlockCopy(value, 0, payload, span.Offset + valueOffset, value.Length);
    }

    private static (int Offset, int Length) FindSpan(byte[] payload, IReadOnlyList<(int Field, int Occurrence)> path,
        int leafWire, IWorkReadOptions limits, CancellationToken token) {
        if (path.Count == 0 || path.Count > limits.MaximumProtobufDepth)
            throw new NotSupportedException("The edit field path exceeds the supported protobuf depth.");
        int start = 0, end = payload.Length;
        for (int level = 0; level < path.Count; level++) {
            int offset = start, count = 0, matches = 0, selectedOffset = -1, selectedLength = 0, selectedWire = -1;
            while (offset < end) {
                token.ThrowIfCancellationRequested();
                if (++count > limits.MaximumProtobufFieldCount) throw IWorkProtobuf.LimitException("Edit field count exceeds the configured limit.");
                ulong key = IWorkProtobuf.ReadVarint(payload, ref offset);
                if (offset > end || key >> 3 is 0 or > 536_870_911) throw new InvalidDataException("Invalid edit field key.");
                int field = (int)(key >> 3), wire = (int)(key & 7), dataOffset = offset, length;
                switch (wire) {
                    case 0:
                        _ = IWorkProtobuf.ReadVarint(payload, ref offset); length = offset - dataOffset; break;
                    case 1: length = 8; offset += length; break;
                    case 2:
                        ulong rawLength = IWorkProtobuf.ReadVarint(payload, ref offset);
                        if (rawLength > int.MaxValue) throw new InvalidDataException("Edit field is too large.");
                        dataOffset = offset; length = (int)rawLength;
                        if (length > end - offset) throw new InvalidDataException("Truncated edit field.");
                        offset += length; break;
                    case 5: length = 4; offset += length; break;
                    default: throw new InvalidDataException("Unsupported edit wire type.");
                }
                if (offset > end) throw new InvalidDataException("Truncated edit field.");
                if (field != path[level].Field) continue;
                int occurrence = matches++;
                if (path[level].Occurrence >= 0 && occurrence != path[level].Occurrence) continue;
                if (selectedOffset >= 0) throw new NotSupportedException("Edit fields must have one unambiguous occurrence.");
                selectedOffset = dataOffset; selectedLength = length; selectedWire = wire;
            }
            bool last = level == path.Count - 1;
            if (selectedOffset < 0 || selectedWire != (last ? leafWire : 2))
                throw new NotSupportedException("The source does not declare the supported fixed-width edit path.");
            if (last) return (selectedOffset, selectedLength);
            start = selectedOffset; end = start + selectedLength;
        }
        throw new NotSupportedException("The source does not declare the supported edit path.");
    }
}
