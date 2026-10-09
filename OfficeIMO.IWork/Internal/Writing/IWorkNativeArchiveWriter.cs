using System.Threading;

namespace OfficeIMO.IWork.Internal;

/// <summary>Writes owned protobuf records and their reference indexes before shared IWA framing.</summary>
internal static class IWorkNativeArchiveWriter {
    internal static byte[] Encode(IEnumerable<(ulong Id, uint Type, IWorkProtoWriter Message)> records,
        int maximumBytes, CancellationToken cancellationToken) {
        IWorkProtoWriter P() => new(maximumBytes, cancellationToken);
        var output = P();
        foreach (var record in records.OrderBy(record => record.Id)) {
            cancellationToken.ThrowIfCancellationRequested();
            byte[] payload = record.Message.ToArray();
            var info = P().UInt(1, record.Type).UInt(3, (ulong)payload.Length);
            foreach (uint value in new uint[] { 1, 0, 5 }) info.UInt(2, value);
            foreach (ulong reference in record.Message.References) info.UInt(5, reference);
            foreach (ulong reference in record.Message.DataReferences) info.UInt(6, reference);
            byte[] header = P().UInt(1, record.Id).Message(2, info).ToArray();
            output.AppendLength(header.Length);
            output.Append(header);
            output.Append(payload);
        }
        return output.ToArray();
    }
}
