using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO;

internal sealed class OfficeBatchConversionReportConverter : JsonConverter<OfficeConversionReport> {
    public override OfficeConversionReport Read(ref Utf8JsonReader reader, Type type, JsonSerializerOptions options) {
        OfficeBatchStoredConversionReport value = JsonSerializer.Deserialize(ref reader,
            OfficeBatchReportSerializer.Context.OfficeBatchStoredConversionReport)
            ?? throw new InvalidDataException("The conversion report is empty.");
        var report = new OfficeConversionReport(value.Source, value.Target, value.OutputPath, value.ItemsWritten, value.Elapsed, value.Diagnostics);
        if (report.HasLoss != value.HasLoss) throw new InvalidDataException("The conversion report loss flag contradicts its findings.");
        return report;
    }
    public override void Write(Utf8JsonWriter writer, OfficeConversionReport value, JsonSerializerOptions options) =>
        JsonSerializer.Serialize(writer, new OfficeBatchStoredConversionReport(value.Source, value.Target, value.OutputPath,
            value.ItemsWritten, value.Elapsed, value.Diagnostics, value.HasLoss), OfficeBatchReportSerializer.Context.OfficeBatchStoredConversionReport);
}

internal sealed record OfficeBatchStoredConversionReport(string Source, string Target, string? OutputPath,
    long ItemsWritten, TimeSpan Elapsed, IReadOnlyList<OfficeDiagnostic> Diagnostics, bool HasLoss);
