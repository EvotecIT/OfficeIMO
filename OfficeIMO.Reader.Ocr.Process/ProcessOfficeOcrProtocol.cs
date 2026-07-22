using OfficeIMO.Reader;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Reader.Ocr.Process;

/// <summary>Portable request file written for an external OCR process.</summary>
public sealed class ProcessOfficeOcrRequest {
    /// <summary>Protocol schema identifier.</summary>
    public string SchemaId { get; set; } = ProcessOfficeOcrProtocol.RequestSchemaId;

    /// <summary>Protocol schema version.</summary>
    public int SchemaVersion { get; set; } = ProcessOfficeOcrProtocol.Version;

    /// <summary>Candidate identifier.</summary>
    public string CandidateId { get; set; } = string.Empty;

    /// <summary>Candidate kind.</summary>
    public string CandidateKind { get; set; } = string.Empty;

    /// <summary>Asset identifier.</summary>
    public string AssetId { get; set; } = string.Empty;

    /// <summary>Asset media type.</summary>
    public string? MediaType { get; set; }

    /// <summary>Source document path or logical name.</summary>
    public string? SourcePath { get; set; }

    /// <summary>Absolute path to the materialized input payload.</summary>
    public string InputPath { get; set; } = string.Empty;

    /// <summary>Absolute path where the process must write an <see cref="OfficeOcrEngineResult"/> JSON object.</summary>
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Requested language expression.</summary>
    public string? Language { get; set; }

    /// <summary>Candidate source location.</summary>
    public ReaderLocation Location { get; set; } = new ReaderLocation();

    /// <summary>Candidate source region.</summary>
    public OfficeDocumentRegion? Region { get; set; }

    /// <summary>Provider-specific scalar options.</summary>
    public IReadOnlyDictionary<string, string> ProviderOptions { get; set; } = new Dictionary<string, string>(StringComparer.Ordinal);
}

/// <summary>Versioned response file written by an external OCR process.</summary>
public sealed class ProcessOfficeOcrResponse {
    /// <summary>Protocol response schema identifier.</summary>
    public string SchemaId { get; set; } = ProcessOfficeOcrProtocol.ResponseSchemaId;

    /// <summary>Protocol schema version.</summary>
    public int SchemaVersion { get; set; } = ProcessOfficeOcrProtocol.Version;

    /// <summary>OCR engine output returned by the external process.</summary>
    public OfficeOcrEngineResult? Result { get; set; }
}

/// <summary>JSON helpers and schema constants for the external OCR process protocol.</summary>
public static class ProcessOfficeOcrProtocol {
    /// <summary>Request schema identifier.</summary>
    public const string RequestSchemaId = "officeimo.reader.ocr.process-request";

    /// <summary>Response schema identifier.</summary>
    public const string ResponseSchemaId = "officeimo.reader.ocr.process-response";

    /// <summary>Current protocol version.</summary>
    public const int Version = 1;

    /// <summary>Serializes a process request using camel-case properties and string enum values.</summary>
    public static string SerializeRequest(ProcessOfficeOcrRequest request, bool indented = false) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        JsonSerializerOptions options = CreateOptions();
        options.WriteIndented = indented;
        var context = new ProcessOfficeOcrJsonSerializerContext(options);
        return JsonSerializer.Serialize(request, context.ProcessOfficeOcrRequest);
    }

    /// <summary>Serializes an engine result suitable for the process response file.</summary>
    public static string SerializeResult(OfficeOcrEngineResult result, bool indented = false) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        JsonSerializerOptions options = CreateOptions();
        options.WriteIndented = indented;
        var context = new ProcessOfficeOcrJsonSerializerContext(options);
        return JsonSerializer.Serialize(new ProcessOfficeOcrResponse { Result = result }, context.ProcessOfficeOcrResponse);
    }

    /// <summary>Deserializes an engine result from the process response file.</summary>
    public static OfficeOcrEngineResult DeserializeResult(string json) {
        if (json == null) throw new ArgumentNullException(nameof(json));
        using (JsonDocument document = JsonDocument.Parse(json)) {
            if (document.RootElement.ValueKind != JsonValueKind.Object) throw new InvalidDataException("OCR process response must be a JSON object.");
            if (!HasProperty(document.RootElement, "schemaId")) throw new InvalidDataException("OCR process response did not contain schemaId.");
            if (!HasProperty(document.RootElement, "schemaVersion")) throw new InvalidDataException("OCR process response did not contain schemaVersion.");
        }
        var context = new ProcessOfficeOcrJsonSerializerContext(CreateOptions());
        ProcessOfficeOcrResponse? response = JsonSerializer.Deserialize(json, context.ProcessOfficeOcrResponse);
        if (response == null) throw new InvalidDataException("OCR process response was empty.");
        if (!string.Equals(response.SchemaId, ResponseSchemaId, StringComparison.Ordinal)) throw new InvalidDataException("OCR process response schema id is not supported.");
        if (response.SchemaVersion != Version) throw new InvalidDataException("OCR process response schema version is not supported.");
        return response.Result ?? throw new InvalidDataException("OCR process response did not contain an engine result.");
    }

    private static bool HasProperty(JsonElement element, string name) {
        foreach (JsonProperty property in element.EnumerateObject()) {
            if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private static JsonSerializerOptions CreateOptions() {
        var options = new JsonSerializerOptions {
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
            PropertyNameCaseInsensitive = true
        };
        return options;
    }
}

[JsonSourceGenerationOptions(
    GenerationMode = JsonSourceGenerationMode.Metadata,
    PropertyNameCaseInsensitive = true,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UseStringEnumConverter = true)]
[JsonSerializable(typeof(ProcessOfficeOcrRequest))]
[JsonSerializable(typeof(ProcessOfficeOcrResponse))]
internal sealed partial class ProcessOfficeOcrJsonSerializerContext : JsonSerializerContext {
}
