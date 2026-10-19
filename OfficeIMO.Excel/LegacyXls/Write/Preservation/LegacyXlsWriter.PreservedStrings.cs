using OfficeIMO.Excel.LegacyXls.Model;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
    internal static (IReadOnlyList<byte[]> Payloads, IReadOnlyList<(int SegmentIndex, int Offset)> Locations)
        EncodePreservedStringAppend(IReadOnlyList<LegacyXlsPreservedString> values, long maximumBytes, CancellationToken cancellationToken, bool includeHeader = false) {
        var writer = new LegacyXlsStringSegmentWriter(maximumBytes);
        if (includeHeader) { writer.WriteUInt32(checked((uint)values.Count)); writer.WriteUInt32(checked((uint)values.Count)); }
        var locations = new List<(int, int)>(values.Count);
        foreach (var value in values) {
            cancellationToken.ThrowIfCancellationRequested();
            var location = writer.WriteUnicodeString(value.Text, value.Runs);
            locations.Add((location.SegmentIndex, location.Offset));
        }
        return (writer.GetPayloads(cancellationToken), locations);
    }
}
