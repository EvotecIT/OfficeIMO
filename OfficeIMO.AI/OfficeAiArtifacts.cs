using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeIMO.Reader;

namespace OfficeIMO.AI;

/// <summary>Portable review artifacts. Saving or reopening an artifact does not approve model-authored content.</summary>
public static class OfficeAiArtifacts {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase, WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };

    /// <summary>Serializes a result with its captured text evidence and image descriptors, without image payloads or credentials.</summary>
    public static string SerializeReport(OfficeAiDocument document, OfficeAiResult result) {
        CheckSource(document, result);
        return JsonSerializer.Serialize(new {
            schema = "officeimo.ai.report.v1",
            source = new { document.SourceHash, document.SourceByteLength, document.PageProvenance, document.Pages },
            evidence = document.Evidence,
            images = document.Images.Select(image => new { image.Id, image.Page, image.MediaType, image.Width, image.Height, image.ByteLength }),
            result
        }, JsonOptions);
    }

    /// <summary>Projects parsed blocks and tables into Reader's transport model, explicitly marked as AI proposals requiring review.</summary>
    public static OfficeDocumentReadResult CreateProposedReadResult(OfficeAiDocument document, OfficeAiResult result) {
        CheckSource(document, result);
        if (result.Operation != OfficeAiOperation.Parse) throw new ArgumentException("Reader structure is produced by Parse operations.", nameof(result));
        var proposed = new OfficeDocumentReadResult {
            Source = new() { SourceHash = document.SourceHash, LengthBytes = document.SourceByteLength },
            CapabilitiesUsed = new[] { "officeimo.ai.parse.proposed" },
            Blocks = result.Blocks.Select(item => item.Block).ToArray(),
            Tables = result.Tables.Select(item => item.Table).ToArray(),
            Diagnostics = new[] { new OfficeDocumentDiagnostic {
                Code = "ai-proposed-requires-review", Message = "Model-authored structure. Validate against the source before use."
            } }
        };
        // Reuse Reader's canonical transport validation and clone its mutable model before returning it.
        return OfficeDocumentReadResultJson.Deserialize(OfficeDocumentReadResultJson.Serialize(proposed));
    }

    private static void CheckSource(OfficeAiDocument document, OfficeAiResult result) {
        ArgumentNullException.ThrowIfNull(document); ArgumentNullException.ThrowIfNull(result);
        if (!string.Equals(document.SourceHash, result.SourceHash, StringComparison.Ordinal))
            throw new ArgumentException("Result and source snapshot fingerprints differ.", nameof(result));
    }
}
