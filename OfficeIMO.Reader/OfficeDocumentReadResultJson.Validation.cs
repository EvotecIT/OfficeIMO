using System;
using System.Collections.Generic;
using System.Text.Json;

namespace OfficeIMO.Reader;

public static partial class OfficeDocumentReadResultJson {
    private static readonly HashSet<string> ReaderInputKindNames = new HashSet<string>(
        Enum.GetNames(typeof(ReaderInputKind)),
        StringComparer.Ordinal);

    private static readonly HashSet<string> DiagnosticSeverityNames = new HashSet<string>(
        Enum.GetNames(typeof(OfficeDocumentDiagnosticSeverity)),
        StringComparer.Ordinal);

    private static readonly HashSet<string> DiagnosticCategoryNames = new HashSet<string>(
        Enum.GetNames(typeof(OfficeDocumentDiagnosticCategory)),
        StringComparer.Ordinal);

    private static readonly string[] RequiredObjectArrayProperties = {
        "chunks",
        "metadata",
        "pages",
        "blocks",
        "tables",
        "assets",
        "links",
        "forms",
        "ocrCandidates",
        "visuals"
    };

    private static readonly HashSet<string> AllowedSourceProperties = new HashSet<string>(new[] {
        "path",
        "sourceId",
        "sourceHash",
        "lastWriteUtc",
        "lengthBytes",
        "title",
        "author",
        "subject",
        "keywords"
    }, StringComparer.Ordinal);

    private static readonly HashSet<string> AllowedDiagnosticProperties = new HashSet<string>(new[] {
        "severity",
        "category",
        "code",
        "message",
        "source",
        "isRecoverable",
        "location",
        "attributes"
    }, StringComparer.Ordinal);

    private static void EnsureNestedTransportContracts(JsonElement root) {
        EnsureEnumString(root.GetProperty("kind"), ReaderInputKindNames, "kind");
        EnsureOptionalString(root, "markdown");
        EnsureOptionalString(root, "html");
        EnsureOptionalString(root, "json");
        EnsureSourceContract(root.GetProperty("source"));
        EnsureStringArray(root.GetProperty("capabilitiesUsed"), "capabilitiesUsed");
        for (int index = 0; index < RequiredObjectArrayProperties.Length; index++) {
            string propertyName = RequiredObjectArrayProperties[index];
            EnsureObjectArray(root.GetProperty(propertyName), propertyName);
        }
        EnsureDiagnosticsArray(root.GetProperty("diagnostics"));
    }

    private static void EnsureSourceContract(JsonElement source) {
        if (source.ValueKind != JsonValueKind.Object) {
            throw new JsonException("Document read result property 'source' must be an object.");
        }
        EnsureKnownNestedProperties(source, AllowedSourceProperties, "source");
        foreach (JsonProperty property in source.EnumerateObject()) {
            bool valid = property.Name == "lengthBytes"
                ? property.Value.ValueKind == JsonValueKind.Number && property.Value.TryGetInt64(out _)
                : property.Value.ValueKind == JsonValueKind.String;
            if (!valid) {
                throw new JsonException($"Document read result property 'source.{property.Name}' has an invalid value.");
            }
        }
    }

    private static void EnsureStringArray(JsonElement array, string propertyName) {
        if (array.ValueKind != JsonValueKind.Array) {
            throw new JsonException($"Document read result property '{propertyName}' must be an array.");
        }
        int index = 0;
        foreach (JsonElement item in array.EnumerateArray()) {
            if (item.ValueKind != JsonValueKind.String) {
                throw new JsonException($"Document read result property '{propertyName}' must contain only strings; item {index} is invalid.");
            }
            index++;
        }
    }

    private static void EnsureObjectArray(JsonElement array, string propertyName) {
        if (array.ValueKind != JsonValueKind.Array) {
            throw new JsonException($"Document read result property '{propertyName}' must be an array.");
        }
        int index = 0;
        foreach (JsonElement item in array.EnumerateArray()) {
            if (item.ValueKind != JsonValueKind.Object) {
                throw new JsonException($"Document read result property '{propertyName}' must contain only objects; item {index} is invalid.");
            }
            index++;
        }
    }

    private static void EnsureDiagnosticsArray(JsonElement diagnostics) {
        EnsureObjectArray(diagnostics, "diagnostics");
        int index = 0;
        foreach (JsonElement diagnostic in diagnostics.EnumerateArray()) {
            EnsureKnownNestedProperties(diagnostic, AllowedDiagnosticProperties, $"diagnostics[{index}]");
            EnsureRequiredDiagnosticString(diagnostic, "severity", index, requireContent: false);
            EnsureRequiredDiagnosticString(diagnostic, "category", index, requireContent: false);
            EnsureEnumString(diagnostic.GetProperty("severity"), DiagnosticSeverityNames, $"diagnostics[{index}].severity");
            EnsureEnumString(diagnostic.GetProperty("category"), DiagnosticCategoryNames, $"diagnostics[{index}].category");
            EnsureRequiredDiagnosticString(diagnostic, "code", index, requireContent: true);
            EnsureRequiredDiagnosticString(diagnostic, "message", index, requireContent: false);
            if (!diagnostic.TryGetProperty("attributes", out JsonElement attributes) || attributes.ValueKind != JsonValueKind.Object) {
                throw new JsonException($"Document diagnostic at index {index} must have an attributes object.");
            }
            foreach (JsonProperty attribute in attributes.EnumerateObject()) {
                if (attribute.Value.ValueKind != JsonValueKind.String) {
                    throw new JsonException($"Document diagnostic at index {index} attribute '{attribute.Name}' must be a string.");
                }
            }
            EnsureOptionalDiagnosticProperty(diagnostic, "source", JsonValueKind.String, index);
            EnsureOptionalDiagnosticProperty(diagnostic, "isRecoverable", JsonValueKind.True, index, JsonValueKind.False);
            EnsureOptionalDiagnosticProperty(diagnostic, "location", JsonValueKind.Object, index);
            index++;
        }
    }

    private static void EnsureOptionalDiagnosticProperty(
        JsonElement diagnostic,
        string propertyName,
        JsonValueKind expected,
        int index,
        JsonValueKind? alternative = null) {
        if (!diagnostic.TryGetProperty(propertyName, out JsonElement value)) return;
        if (value.ValueKind != expected && (!alternative.HasValue || value.ValueKind != alternative.Value)) {
            throw new JsonException($"Document diagnostic at index {index} property '{propertyName}' has an invalid value.");
        }
    }

    private static void EnsureRequiredDiagnosticString(JsonElement diagnostic, string propertyName, int index, bool requireContent) {
        if (!diagnostic.TryGetProperty(propertyName, out JsonElement value) || value.ValueKind != JsonValueKind.String) {
            throw new JsonException($"Document diagnostic at index {index} must have a {propertyName} string.");
        }
        if (requireContent && string.IsNullOrWhiteSpace(value.GetString())) {
            throw new JsonException($"Document diagnostic at index {index} must have a non-empty {propertyName}.");
        }
    }

    private static void EnsureOptionalString(JsonElement value, string propertyName) {
        if (!value.TryGetProperty(propertyName, out JsonElement property)) return;
        if (property.ValueKind != JsonValueKind.String) {
            throw new JsonException($"Document read result property '{propertyName}' must be a string when present.");
        }
    }

    private static void EnsureEnumString(JsonElement value, HashSet<string> allowedValues, string propertyName) {
        if (value.ValueKind != JsonValueKind.String || !allowedValues.Contains(value.GetString() ?? string.Empty)) {
            throw new JsonException($"Document read result property '{propertyName}' has an invalid enum value.");
        }
    }

    private static void EnsureKnownNestedProperties(JsonElement value, HashSet<string> allowed, string context) {
        foreach (JsonProperty property in value.EnumerateObject()) {
            if (!allowed.Contains(property.Name)) {
                throw new JsonException($"Unknown document read result property '{context}.{property.Name}'.");
            }
        }
    }

    private static void EnsureStringCollection(IReadOnlyList<string>? values, string propertyName) {
        if (values == null) return;
        for (int index = 0; index < values.Count; index++) {
            if (values[index] == null) {
                throw new JsonException($"Document read result property '{propertyName}' contains a null item at index {index}.");
            }
        }
    }
}
