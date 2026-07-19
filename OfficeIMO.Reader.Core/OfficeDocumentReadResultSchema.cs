using System;
using System.IO;
using System.Reflection;

namespace OfficeIMO.Reader;

/// <summary>
/// Stable schema discovery and compatibility helpers for document read result transport payloads.
/// </summary>
public static partial class OfficeDocumentReadResultSchema {
    /// <summary>
    /// First schema version covered by the stable compatibility contract.
    /// Versions 1 through 4 were experimental and are not accepted by the transport reader.
    /// </summary>
    public const int MinimumSupportedVersion = 5;

    /// <summary>
    /// Current schema version emitted and accepted by this package.
    /// </summary>
    public const int CurrentVersion = 6;

    /// <summary>
    /// Stable JSON Schema identifier for current version 6 payloads.
    /// </summary>
    public const string JsonSchemaId = "urn:officeimo:schema:document-read-result:6";

    /// <summary>
    /// File name used for the packaged current JSON Schema artifact.
    /// </summary>
    public const string JsonSchemaFileName = "officeimo.document.read-result.v6.schema.json";

    /// <summary>
    /// Returns true when a schema header can be consumed by this package.
    /// </summary>
    public static bool IsSupported(string? schemaId, int schemaVersion) {
        return string.Equals(schemaId, Id, StringComparison.Ordinal) &&
               schemaVersion >= MinimumSupportedVersion &&
               schemaVersion <= CurrentVersion;
    }

    /// <summary>
    /// Throws when a schema header cannot be consumed by this package.
    /// </summary>
    public static void EnsureSupported(string? schemaId, int schemaVersion) {
        if (IsSupported(schemaId, schemaVersion)) return;
        throw new OfficeDocumentReadResultSchemaException(schemaId, schemaVersion);
    }

    /// <summary>
    /// Loads the JSON Schema artifact embedded in the assembly and included in the NuGet package.
    /// </summary>
    public static string GetJsonSchema() => GetJsonSchema(CurrentVersion);

    /// <summary>
    /// Loads one supported versioned JSON Schema artifact embedded in the assembly and included in the NuGet package.
    /// </summary>
    public static string GetJsonSchema(int schemaVersion) {
        EnsureSupported(Id, schemaVersion);
        string fileName = $"officeimo.document.read-result.v{schemaVersion}.schema.json";
        string resourceName = "OfficeIMO.Reader.Schemas." + fileName;
        Assembly assembly = typeof(OfficeDocumentReadResultSchema).Assembly;
        using Stream? stream = assembly.GetManifestResourceStream(resourceName);
        if (stream == null) {
            throw new InvalidOperationException($"Embedded JSON Schema resource '{resourceName}' was not found.");
        }
        using var reader = new StreamReader(stream);
        return reader.ReadToEnd();
    }
}

/// <summary>
/// Describes an unsupported or invalid OfficeIMO document read result schema header.
/// </summary>
public sealed class OfficeDocumentReadResultSchemaException : Exception {
    /// <summary>
    /// Creates a schema compatibility exception.
    /// </summary>
    public OfficeDocumentReadResultSchemaException(string? schemaId, int schemaVersion)
        : base(CreateMessage(schemaId, schemaVersion)) {
        SchemaId = schemaId;
        SchemaVersion = schemaVersion;
    }

    /// <summary>
    /// Schema identifier found in the payload.
    /// </summary>
    public string? SchemaId { get; }

    /// <summary>
    /// Schema version found in the payload.
    /// </summary>
    public int SchemaVersion { get; }

    private static string CreateMessage(string? schemaId, int schemaVersion) {
        string displayId = string.IsNullOrWhiteSpace(schemaId) ? "<missing>" : schemaId!;
        return $"Document read result schema '{displayId}' version {schemaVersion} is not supported. " +
               $"Expected '{OfficeDocumentReadResultSchema.Id}' version " +
               $"{OfficeDocumentReadResultSchema.MinimumSupportedVersion} through {OfficeDocumentReadResultSchema.CurrentVersion}.";
    }
}
