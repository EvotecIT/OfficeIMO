using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Patches same-length payloads while retaining complete native ArchiveInfo/MessageInfo envelopes and siblings.</summary>
internal static class IWorkPreservingArchiveEdit {
    internal static byte[] Rewrite(byte[] raw, string path, IReadOnlyDictionary<IWorkArchiveRecord, byte[]> changes,
        IWorkReadOptions limits, CancellationToken token) {
        var selected = changes.Where(pair => pair.Key.EntryPath == path).ToDictionary(pair => pair.Key.Identifier);
        var found = new HashSet<ulong>(); int offset = 0;
        while (offset < raw.Length) {
            token.ThrowIfCancellationRequested();
            ulong length = IWorkProtobuf.ReadVarint(raw, ref offset);
            if (length > (ulong)limits.MaximumArchiveInfoBytes || length > (ulong)(raw.Length - offset))
                throw new InvalidDataException("Invalid preserved ArchiveInfo length.");
            var header = new byte[(int)length]; Buffer.BlockCopy(raw, offset, header, 0, header.Length); offset += header.Length;
            IWorkWireMessage info = IWorkProtobuf.Parse(header, limits);
            ulong id = info.GetUnsigned(1) ?? throw new InvalidDataException("Preserved ArchiveInfo lacks an identifier.");
            var messages = info.GetRepeatedMessages(2);
            for (int index = 0; index < messages.Count; index++) {
                token.ThrowIfCancellationRequested();
                ulong rawLength = messages[index].GetUnsigned(3) ?? throw new InvalidDataException("Preserved MessageInfo lacks a payload length.");
                if (rawLength > (ulong)(raw.Length - offset)) throw new InvalidDataException("Truncated preserved payload.");
                int payloadLength = (int)rawLength;
                if (selected.TryGetValue(id, out var change) && change.Key.PayloadIndex == index) {
                    if (change.Value.Length != payloadLength || change.Key.PayloadLength != payloadLength || !found.Add(id))
                        throw new InvalidDataException("Preserving edits must match one same-length source payload.");
                    for (int position = 0; position < payloadLength; position++) {
                        if ((position & 65_535) == 0) token.ThrowIfCancellationRequested();
                        if (raw[offset + position] != change.Key.Payload[position])
                            throw new InvalidDataException("The preserved archive differs from the captured source record.");
                    }
                    Buffer.BlockCopy(change.Value, 0, raw, offset, payloadLength);
                }
                offset += payloadLength;
            }
        }
        if (found.Count != selected.Count) throw new InvalidDataException("A requested edit was not found in its source archive.");
        return raw;
    }
}
