using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Workflows;

/// <summary>Shared, source-generated JSON for bounded batch reports and streaming item/summary events.</summary>
public static class OfficeConversionBatchReportSerializer {
    /// <summary>Maximum UTF-8 size of a report or report event.</summary>
    public const int MaximumReportBytes = 20 * 1024 * 1024;
    internal static readonly OfficeConversionReportJsonContext Context = new(new JsonSerializerOptions {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase, UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        MaxDepth = 32, Converters = {
            new JsonStringEnumConverter<OfficeWorkflowStatus>(allowIntegerValues: false),
            new JsonStringEnumConverter<OfficeWorkflowDiagnosticSeverity>(allowIntegerValues: false),
            new JsonStringEnumConverter<OfficeConversionJobRunState>(allowIntegerValues: false)
        }
    });

    /// <summary>Serializes a bounded report using the same contract on every host.</summary>
    public static string Serialize(OfficeConversionBatchReport report) {
        Validate(report);
        return SerializeBounded(report, Context.OfficeConversionBatchReport);
    }

    /// <summary>Reads strict report JSON, rejecting unknown fields, duplicate properties, invalid enums and excess observations.</summary>
    public static OfficeConversionBatchReport Deserialize(string json) {
        OfficeConversionBatchReport report = DeserializeBounded(json, Context.OfficeConversionBatchReport);
        Validate(report);
        return report;
    }

    /// <summary>Serializes one NDJSON item event; oversized individual records are rejected.</summary>
    public static string SerializeItemEvent(OfficeConversionBatchItemResult item) {
        ArgumentNullException.ThrowIfNull(item);
        if (!TryGetItemSize(item, OfficeConversionBatchReportCollector.MaximumItemBytes, out _))
            throw new InvalidDataException("The batch item exceeds its observation limit.");
        return SerializeBounded(new OfficeConversionBatchItemEvent(OfficeConversionBatchReport.CurrentSchema, "item", item), Context.OfficeConversionBatchItemEvent);
    }

    /// <summary>Serializes one NDJSON summary event. Hosts may reflect an additional requested-delivery failure in succeeded.</summary>
    public static string SerializeSummaryEvent(OfficeConversionBatchResult summary, bool succeeded) {
        ValidateSummary(summary);
        return SerializeBounded(new OfficeConversionBatchSummaryEvent(OfficeConversionBatchReport.CurrentSchema, "summary", succeeded, summary), Context.OfficeConversionBatchSummaryEvent);
    }

    internal static void Validate(OfficeConversionBatchReport report) {
        ArgumentNullException.ThrowIfNull(report);
        if (report.Schema != OfficeConversionBatchReport.CurrentSchema || report.Type != "report")
            throw new NotSupportedException("Unsupported batch report contract.");
        if (report.Summary == null || report.Items == null || report.Items.Count > OfficeConversionBatchReportCollector.MaximumItems)
            throw new InvalidDataException("The batch report summary or observations are invalid.");
        var summary = report.Summary;
        ValidateSummary(summary);
        if (report.Succeeded && (summary.Cancelled || summary.Failed != 0 || summary.FailureSummary != null))
            throw new InvalidDataException("A stopped or failed batch cannot be reported as successful.");
        int remaining = OfficeConversionBatchReportCollector.MaximumItemBytes;
        foreach (var item in report.Items) {
            if (item == null || !TryGetItemSize(item, remaining, out int bytes))
                throw new InvalidDataException("The batch report observations exceed their byte limit.");
            remaining -= bytes;
        }
    }

    private static void ValidateSummary(OfficeConversionBatchResult summary) {
        ArgumentNullException.ThrowIfNull(summary);
        if (summary.Selected < 0 || summary.Completed < 0 || summary.Reused < 0 || summary.Failed < 0 || summary.Skipped < 0 ||
            summary.Completed > summary.Selected || summary.Reused > summary.Completed || summary.Failed > summary.Selected - summary.Completed)
            throw new InvalidDataException("The batch report counts are invalid.");
        if (summary.FailureSummary?.Length > 4096) throw new InvalidDataException("The batch failure summary is oversized.");
    }

    internal static void CheckJson(string json) {
        ArgumentNullException.ThrowIfNull(json);
        if (Encoding.UTF8.GetByteCount(json) > MaximumReportBytes) throw new InvalidDataException("The report exceeds its 20 MiB limit.");
        using var document = JsonDocument.Parse(json, new JsonDocumentOptions { MaxDepth = 32 });
        OfficeConversionJobSerializer.RejectDuplicates(document.RootElement);
    }

    internal static T DeserializeBounded<T>(string json, System.Text.Json.Serialization.Metadata.JsonTypeInfo<T> type) where T : class {
        CheckJson(json);
        try {
            return JsonSerializer.Deserialize(json, type) ?? throw new InvalidDataException("The report is empty.");
        } catch (ArgumentException error) {
            throw new InvalidDataException("The report contains invalid object values.", error);
        }
    }

    internal static string SerializeBounded<T>(T value, System.Text.Json.Serialization.Metadata.JsonTypeInfo<T> type) {
        using var stream = new BoundedReportStream(MaximumReportBytes);
        JsonSerializer.Serialize(stream, value, type);
        return Encoding.UTF8.GetString(stream.ToArray());
    }

    internal static bool TryGetItemSize(OfficeConversionBatchItemResult item, int maximum, out int bytes) {
        bytes = 0;
        // Reject a single huge string before the JSON writer requests a correspondingly huge encoding buffer.
        if (maximum < 1 || item.InputPath == null || item.Summary == null || item.Diagnostics == null ||
            item.InputPath.Length > maximum || item.OutputPath?.Length > maximum || item.Summary.Length > maximum) return false;
        long characters = item.InputPath.Length + (long)(item.OutputPath?.Length ?? 0) + item.Summary.Length;
        foreach (var diagnostic in item.Diagnostics) {
            if (diagnostic == null) return false;
            characters += diagnostic.Code.Length + (long)diagnostic.Message.Length + (diagnostic.Stage?.Length ?? 0);
            foreach (var detail in diagnostic.Details) characters += detail.Key.Length + (long)(detail.Value?.Length ?? 0);
            if (characters > maximum) return false;
        }
        if (characters > maximum || item.Status is { } status && !Enum.IsDefined(status)) return false;
        try {
            using var stream = new BoundedReportStream(maximum);
            JsonSerializer.Serialize(stream, item, Context.OfficeConversionBatchItemResult);
            bytes = checked((int)stream.Length);
            return true;
        } catch (InvalidDataException) { return false; }
    }
}

internal sealed record OfficeConversionBatchItemEvent(string Schema, string Type, OfficeConversionBatchItemResult Item);
internal sealed record OfficeConversionBatchSummaryEvent(string Schema, string Type, bool Succeeded, OfficeConversionBatchResult Summary);

internal sealed class BoundedReportStream(int maximum) : MemoryStream {
    public override void Write(byte[] buffer, int offset, int count) {
        Check(count); base.Write(buffer, offset, count);
    }
    public override void Write(ReadOnlySpan<byte> buffer) { Check(buffer.Length); base.Write(buffer); }
    private void Check(int count) {
        if (count > maximum - Length) throw new InvalidDataException("The report exceeds its byte limit.");
    }
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, GenerationMode = JsonSourceGenerationMode.Metadata,
    UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow)]
[JsonSerializable(typeof(OfficeConversionBatchReport))]
[JsonSerializable(typeof(OfficeConversionBatchItemEvent))]
[JsonSerializable(typeof(OfficeConversionBatchSummaryEvent))]
[JsonSerializable(typeof(OfficeConversionBatchItemResult))]
[JsonSerializable(typeof(OfficeConversionJobRunReport))]
internal sealed partial class OfficeConversionReportJsonContext : JsonSerializerContext;
