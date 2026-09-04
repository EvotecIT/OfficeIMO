using OfficeIMO.Ocr;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Ocr.Process;

/// <summary>Portable request file written for an external OCR process.</summary>
public sealed class ProcessOcrRequest {
    /// <summary>Protocol schema identifier.</summary>
    public string SchemaId { get; set; } = ProcessOcrProtocol.RequestSchemaId;
    /// <summary>Protocol schema version.</summary>
    public int SchemaVersion { get; set; } = ProcessOcrProtocol.Version;
    /// <summary>Caller-owned recognition candidate identifier.</summary>
    public string? CandidateId { get; set; }
    /// <summary>Caller-defined candidate kind.</summary>
    public string? CandidateKind { get; set; }
    /// <summary>Input media type.</summary>
    public string MediaType { get; set; } = string.Empty;
    /// <summary>Original or synthetic file name.</summary>
    public string? FileName { get; set; }
    /// <summary>Stable source artifact identifier.</summary>
    public string? SourceId { get; set; }
    /// <summary>Source path or logical name.</summary>
    public string? SourceName { get; set; }
    /// <summary>One-based source page or frame number.</summary>
    public int? PageNumber { get; set; }
    /// <summary>Raster width in pixels.</summary>
    public int? PixelWidth { get; set; }
    /// <summary>Raster height in pixels.</summary>
    public int? PixelHeight { get; set; }
    /// <summary>Absolute path to the materialized input payload.</summary>
    public string InputPath { get; set; } = string.Empty;
    /// <summary>Absolute path where the process must write an <see cref="OcrResult"/> response.</summary>
    public string OutputPath { get; set; } = string.Empty;
    /// <summary>Requested language expression.</summary>
    public string? Language { get; set; }
    /// <summary>Source region represented by the payload.</summary>
    public OcrRegion? Region { get; set; }
    /// <summary>Coordinate unit used by <see cref="Region"/>.</summary>
    public OcrCoordinateUnit RegionCoordinateUnit { get; set; } = OcrCoordinateUnit.Pixels;
    /// <summary>Provider-specific scalar options.</summary>
    public IReadOnlyDictionary<string, string> ProviderOptions { get; set; } =
        new Dictionary<string, string>(StringComparer.Ordinal);
}

/// <summary>Versioned response file written by an external OCR process.</summary>
public sealed class ProcessOcrResponse {
    /// <summary>Protocol response schema identifier.</summary>
    public string SchemaId { get; set; } = ProcessOcrProtocol.ResponseSchemaId;
    /// <summary>Protocol schema version.</summary>
    public int SchemaVersion { get; set; } = ProcessOcrProtocol.Version;
    /// <summary>OCR engine output returned by the external process.</summary>
    public OcrResult? Result { get; set; }
}

/// <summary>JSON helpers and schema constants for the external OCR process protocol.</summary>
public static class ProcessOcrProtocol {
    /// <summary>Request schema identifier.</summary>
    public const string RequestSchemaId = "officeimo.ocr.process-request";
    /// <summary>Response schema identifier.</summary>
    public const string ResponseSchemaId = "officeimo.ocr.process-response";
    /// <summary>Current protocol version.</summary>
    public const int Version = 2;

    /// <summary>Serializes a process request using camel-case properties and string enum values.</summary>
    public static string SerializeRequest(ProcessOcrRequest request, bool indented = false) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        JsonSerializerOptions options = CreateOptions();
        options.WriteIndented = indented;
        var context = new ProcessOcrJsonSerializerContext(options);
        return JsonSerializer.Serialize(request, context.ProcessOcrRequest);
    }

    /// <summary>Serializes an engine result suitable for the process response file.</summary>
    public static string SerializeResult(OcrResult result, bool indented = false) {
        if (result == null) throw new ArgumentNullException(nameof(result));
        JsonSerializerOptions options = CreateOptions();
        options.WriteIndented = indented;
        var context = new ProcessOcrJsonSerializerContext(options);
        return JsonSerializer.Serialize(new ProcessOcrResponse { Result = result }, context.ProcessOcrResponse);
    }

    /// <summary>Deserializes an engine result from the process response file.</summary>
    public static OcrResult DeserializeResult(string json) {
        if (json == null) throw new ArgumentNullException(nameof(json));
        using (JsonDocument document = JsonDocument.Parse(json)) {
            if (document.RootElement.ValueKind != JsonValueKind.Object) throw new InvalidDataException("OCR process response must be a JSON object.");
            if (!HasProperty(document.RootElement, "schemaId")) throw new InvalidDataException("OCR process response did not contain schemaId.");
            if (!HasProperty(document.RootElement, "schemaVersion")) throw new InvalidDataException("OCR process response did not contain schemaVersion.");
        }
        var context = new ProcessOcrJsonSerializerContext(CreateOptions());
        ProcessOcrResponse? response = JsonSerializer.Deserialize(json, context.ProcessOcrResponse);
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

    private static JsonSerializerOptions CreateOptions() => new JsonSerializerOptions {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true
    };
}

[JsonSourceGenerationOptions(
    GenerationMode = JsonSourceGenerationMode.Metadata,
    PropertyNameCaseInsensitive = true,
    PropertyNamingPolicy = JsonKnownNamingPolicy.CamelCase,
    UseStringEnumConverter = true)]
[JsonSerializable(typeof(ProcessOcrRequest))]
[JsonSerializable(typeof(ProcessOcrResponse))]
internal sealed partial class ProcessOcrJsonSerializerContext : JsonSerializerContext {
}
