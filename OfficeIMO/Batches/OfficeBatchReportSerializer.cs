using OfficeIMO.Internal;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO;

/// <summary>Shared, source-generated JSON for bounded batch reports and streaming item/summary events.</summary>
internal static class OfficeBatchReportSerializer {
    /// <summary>Maximum UTF-8 size of a report or report event.</summary>
    public const int MaximumReportBytes = 20 * 1024 * 1024;
    internal static readonly OfficeBatchReportJsonContext Context = new(new JsonSerializerOptions {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase, UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
        MaxDepth = 32, Converters = {
            new JsonStringEnumConverter<OfficeBatchItemStatus>(allowIntegerValues: false),
            new JsonStringEnumConverter<OfficeDiagnosticSeverity>(allowIntegerValues: false),
            new JsonStringEnumConverter<OfficeConversionLossKind>(allowIntegerValues: false),
            new JsonStringEnumConverter<OfficeBatchRunState>(allowIntegerValues: false)
        }
    });

    /// <summary>Serializes a bounded report using the same contract on every host.</summary>
    public static string Serialize(OfficeBatchReport report) {
        Validate(report);
        return SerializeBounded(report, Context.OfficeBatchReport);
    }

    /// <summary>Reads strict report JSON, rejecting unknown fields, duplicate properties, invalid enums and excess observations.</summary>
    public static OfficeBatchReport Deserialize(string json) {
        OfficeBatchReport report = DeserializeBounded(json, Context.OfficeBatchReport);
        Validate(report);
        return report;
    }

    /// <summary>Serializes one NDJSON item event; oversized individual records are rejected.</summary>
    public static string SerializeItemEvent(OfficeBatchItemReport item) {
        if (item == null) throw new ArgumentNullException(nameof(item));
        if (!TryGetItemSize(item, OfficeBatchReportCollector.MaximumItemBytes, out _))
            throw new InvalidDataException("The batch item exceeds its observation limit.");
        return SerializeBounded(new OfficeBatchItemEvent(OfficeBatchReport.CurrentSchema, "item", item), Context.OfficeBatchItemEvent);
    }

    /// <summary>Serializes one NDJSON summary event. Hosts may reflect an additional requested-delivery failure in succeeded.</summary>
    public static string SerializeSummaryEvent(OfficeBatchSummary summary, bool succeeded) {
        ValidateSummary(summary);
        return SerializeBounded(new OfficeBatchSummaryEvent(OfficeBatchReport.CurrentSchema, "summary", succeeded, summary), Context.OfficeBatchSummaryEvent);
    }

    internal static void Validate(OfficeBatchReport report) {
        if (report == null) throw new ArgumentNullException(nameof(report));
        if (report.Schema != OfficeBatchReport.CurrentSchema)
            throw new NotSupportedException("Unsupported batch report contract.");
        if (report.Summary == null || report.Items == null || report.Items.Count > OfficeBatchReportCollector.MaximumItems)
            throw new InvalidDataException("The batch report summary or observations are invalid.");
        var summary = report.Summary;
        ValidateSummary(summary);
        if (report.Succeeded && (summary.Cancelled || summary.Failed != 0 || summary.FailureSummary != null))
            throw new InvalidDataException("A stopped or failed batch cannot be reported as successful.");
        int remaining = OfficeBatchReportCollector.MaximumItemBytes;
        foreach (var item in report.Items) {
            if (item == null || !TryGetItemSize(item, remaining, out int bytes))
                throw new InvalidDataException("The batch report observations exceed their byte limit.");
            remaining -= bytes;
        }
    }

    private static void ValidateSummary(OfficeBatchSummary summary) {
        if (summary == null) throw new ArgumentNullException(nameof(summary));
        if (summary.Selected < 0 || summary.Completed < 0 || summary.Reused < 0 || summary.Failed < 0 || summary.Skipped < 0 ||
            summary.Completed > summary.Selected || summary.Reused > summary.Completed || summary.Failed > summary.Selected - summary.Completed)
            throw new InvalidDataException("The batch report counts are invalid.");
        if (summary.FailureSummary?.Length > 4096) throw new InvalidDataException("The batch failure summary is oversized.");
    }

    internal static void CheckJson(string json) {
        if (json == null) throw new ArgumentNullException(nameof(json));
        if (Encoding.UTF8.GetByteCount(json) > MaximumReportBytes) throw new InvalidDataException("The report exceeds its 20 MiB limit.");
        using var document = JsonDocument.Parse(json, new JsonDocumentOptions { MaxDepth = 32 });
        OfficeBatchJson.RejectDuplicates(document.RootElement);
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

    internal static bool TryGetItemSize(OfficeBatchItemReport item, int maximum, out int bytes) {
        bytes = 0;
        // Reject a single huge string before the JSON writer requests a correspondingly huge encoding buffer.
        if (maximum < 1 || item.InputPath == null || item.Summary == null || item.Diagnostics == null ||
            item.InputPath.Length > maximum || item.OutputPath?.Length > maximum || item.Summary.Length > maximum) return false;
        long characters = item.InputPath.Length + (long)(item.OutputPath?.Length ?? 0) + item.Summary.Length;
        foreach (var diagnostic in item.Diagnostics) {
            if (diagnostic == null) return false;
            characters += diagnostic.Code.Length + (long)diagnostic.Message.Length + (diagnostic.Location?.Length ?? 0) + (diagnostic.Format?.Length ?? 0);
            if (characters > maximum) return false;
        }
        if (item.Report is { } report) {
            characters += report.Source.Length + (long)report.Target.Length + (report.OutputPath?.Length ?? 0);
            foreach (OfficeDiagnostic diagnostic in report.Diagnostics) {
                characters += diagnostic.Code.Length + (long)diagnostic.Message.Length + (diagnostic.Location?.Length ?? 0) + (diagnostic.Format?.Length ?? 0);
                if (characters > maximum) return false;
            }
        }
        if (characters > maximum || item.Status is { } status && !OfficeRuntimeCompatibility.IsDefined(status)) return false;
        try {
            using var stream = new BoundedReportStream(maximum);
            JsonSerializer.Serialize(stream, item, Context.OfficeBatchItemReport);
            bytes = checked((int)stream.Length);
            return true;
        } catch (InvalidDataException) { return false; }
    }
}

internal sealed record OfficeBatchItemEvent(string Schema, string Type, OfficeBatchItemReport Item);
internal sealed record OfficeBatchSummaryEvent(string Schema, string Type, bool Succeeded, OfficeBatchSummary Summary);

internal sealed class BoundedReportStream(int maximum) : MemoryStream {
    public override void Write(byte[] buffer, int offset, int count) {
        Check(count); base.Write(buffer, offset, count);
    }
#if NET8_0_OR_GREATER
    public override void Write(ReadOnlySpan<byte> buffer) { Check(buffer.Length); base.Write(buffer); }
#endif
    private void Check(int count) {
        if (count > maximum - Length) throw new InvalidDataException("The report exceeds its byte limit.");
    }
}

[JsonSourceGenerationOptions(PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase, GenerationMode = JsonSourceGenerationMode.Metadata,
    UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow)]
[JsonSerializable(typeof(OfficeBatchReport))]
[JsonSerializable(typeof(OfficeBatchItemEvent))]
[JsonSerializable(typeof(OfficeBatchSummaryEvent))]
[JsonSerializable(typeof(OfficeBatchItemReport))]
[JsonSerializable(typeof(OfficeBatchStoredConversionReport))]
internal sealed partial class OfficeBatchReportJsonContext : JsonSerializerContext;
