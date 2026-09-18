using System.Globalization;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PdfFormatConversionRecordManifest {
    internal const int RecordCount = 120;
    internal const string Description = "Deterministic account conversion evidence";

    internal static IReadOnlyList<string> CreateRequiredText(string heading, int recordCount = RecordCount, string? structureMarker = null) {
        if (recordCount < 1) {
            throw new ArgumentOutOfRangeException(nameof(recordCount));
        }

        var required = new List<string>(3 + (recordCount * 4)) {
            heading,
            Description
        };
        if (!string.IsNullOrWhiteSpace(structureMarker)) {
            required.Add(structureMarker!);
        }
        for (int index = 1; index <= recordCount; index++) {
            required.Add(RecordMarker(index));
            required.Add(CustomerMarker(index));
            required.Add(AmountMarker(index));
            required.Add(StatusMarker(index));
        }
        return required;
    }

    internal static string RecordLine(int index) =>
        $"{RecordMarker(index)} {CustomerMarker(index)} {Description} {AmountMarker(index)} {StatusMarker(index)}";

    internal static string RecordMarker(int index) => $"RECORD-{index:D4}";

    internal static string CustomerMarker(int index) => $"CUSTOMER-{index:D4}";

    internal static string AmountMarker(int index) =>
        "AMOUNT-" + (index * 37.25M).ToString("0.00", CultureInfo.InvariantCulture);

    internal static string StatusMarker(int index) =>
        (index % 7 == 0 ? "STATUS-REVIEW-" : "STATUS-APPROVED-") + index.ToString("D4", CultureInfo.InvariantCulture);
}
