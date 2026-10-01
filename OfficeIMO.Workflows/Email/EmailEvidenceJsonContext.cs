using System.Text.Json.Serialization;

namespace OfficeIMO.Workflows;

[JsonSourceGenerationOptions(WriteIndented = true)]
[JsonSerializable(typeof(EmailEvidenceManifest))]
internal sealed partial class EmailEvidenceJsonContext : JsonSerializerContext;
