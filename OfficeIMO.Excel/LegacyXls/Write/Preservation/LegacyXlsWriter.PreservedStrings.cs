using OfficeIMO.Excel.LegacyXls.Model;
using System.Threading;

namespace OfficeIMO.Excel.LegacyXls.Write;

internal static partial class LegacyXlsWriter {
    internal static (IReadOnlyList<byte[]> Payloads, IReadOnlyList<(int SegmentIndex, int Offset)> Locations)
        EncodePreservedStringAppend(IReadOnlyList<string> values, long maximumBytes, CancellationToken cancellationToken) {
        var writer = new LegacyXlsStringSegmentWriter(maximumBytes);
        var locations = new List<(int, int)>(values.Count);
        foreach (string value in values) {
            cancellationToken.ThrowIfCancellationRequested();
            var location = writer.WriteUnicodeString(value, Array.Empty<LegacyXlsTextFormattingRun>());
            locations.Add((location.SegmentIndex, location.Offset));
        }
        return (writer.GetPayloads(cancellationToken), locations);
    }
}
