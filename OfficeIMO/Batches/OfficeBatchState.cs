using System.Text.Json.Serialization;

namespace OfficeIMO;

internal sealed record OfficeBatchCheckpointPlan(string Schema, string ConfigurationSha256);
internal sealed record OfficeBatchReceipt(string Schema, string ConfigurationSha256, string InputSha256,
    string? OutputSha256, OfficeBatchItemStatus Status, string Summary, OfficeBatchStoredDiagnostic[] Diagnostics, int DiagnosticCount,
    string? PendingStageId = null);
internal sealed record OfficeBatchStoredDiagnostic(string Code, string Message, OfficeDiagnosticSeverity Severity,
    string? Stage, string? Source, string? LossKind, string? Location = null);

[JsonSourceGenerationOptions(UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow)]
[JsonSerializable(typeof(OfficeBatchSummary))]
[JsonSerializable(typeof(OfficeBatchCheckpointPlan))]
[JsonSerializable(typeof(OfficeBatchReceipt))]
internal partial class OfficeBatchStateJsonContext : JsonSerializerContext { }
