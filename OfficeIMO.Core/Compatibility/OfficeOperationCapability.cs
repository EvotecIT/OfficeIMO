using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO;

/// <summary>Identifies a package-neutral document operation.</summary>
public enum OfficeOperationKind {
    /// <summary>Create a new artifact.</summary>
    Create,
    /// <summary>Read an artifact into an OfficeIMO model.</summary>
    Read,
    /// <summary>Edit or rewrite an artifact.</summary>
    Edit,
    /// <summary>Retain source content that is not projected into an editable model.</summary>
    Preserve,
    /// <summary>Inspect structure, metadata, protection, or provenance.</summary>
    Inspect,
    /// <summary>Validate a format, protection, or provenance contract.</summary>
    Validate,
    /// <summary>Remove selected content through a format-aware operation.</summary>
    Remove,
    /// <summary>Convert between document formats or semantic models.</summary>
    Convert,
    /// <summary>Export a rendered or derived representation.</summary>
    Export
}

/// <summary>Describes the supported outcome of one package-neutral operation.</summary>
public enum OfficeOperationSupportState {
    /// <summary>The operation is implemented and backed by named evidence.</summary>
    Supported,
    /// <summary>A bounded subset is implemented and the limitations are explicit.</summary>
    Partial,
    /// <summary>Source content is retained without claiming editable interpretation.</summary>
    Preserved,
    /// <summary>The operation deliberately fails to prevent unsafe or misleading output.</summary>
    Rejected,
    /// <summary>The operation is not implemented.</summary>
    Unsupported,
    /// <summary>The operation does not apply to this capability.</summary>
    NotApplicable
}

/// <summary>One operation row projected from an owning OfficeIMO capability contract.</summary>
public sealed class OfficeOperationCapability {
    /// <summary>Creates one package-neutral operation row.</summary>
    public OfficeOperationCapability(
        string id,
        string packageId,
        string formatId,
        string capabilityId,
        OfficeOperationKind operation,
        OfficeOperationSupportState state,
        string publicApi,
        string evidence,
        string sourceCatalog,
        IEnumerable<string>? extensions = null,
        string? targetFormatId = null,
        string? limitation = null) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Operation id cannot be empty.", nameof(id));
        if (string.IsNullOrWhiteSpace(packageId)) throw new ArgumentException("Package id cannot be empty.", nameof(packageId));
        if (string.IsNullOrWhiteSpace(formatId)) throw new ArgumentException("Format id cannot be empty.", nameof(formatId));
        if (string.IsNullOrWhiteSpace(capabilityId)) throw new ArgumentException("Capability id cannot be empty.", nameof(capabilityId));
        if (string.IsNullOrWhiteSpace(publicApi)) throw new ArgumentException("Public API cannot be empty.", nameof(publicApi));
        if (string.IsNullOrWhiteSpace(evidence)) throw new ArgumentException("Evidence cannot be empty.", nameof(evidence));
        if (string.IsNullOrWhiteSpace(sourceCatalog)) throw new ArgumentException("Source catalog cannot be empty.", nameof(sourceCatalog));

        string[] normalizedExtensions = (extensions ?? Array.Empty<string>())
            .Select(NormalizeExtension)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(static value => value, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        Id = id.Trim();
        PackageId = packageId.Trim();
        FormatId = formatId.Trim();
        TargetFormatId = targetFormatId?.Trim() ?? string.Empty;
        CapabilityId = capabilityId.Trim();
        Operation = operation;
        State = state;
        PublicApi = publicApi.Trim();
        Evidence = evidence.Trim();
        SourceCatalog = sourceCatalog.Trim();
        Extensions = new ReadOnlyCollection<string>(normalizedExtensions);
        Limitation = limitation?.Trim() ?? string.Empty;
    }

    /// <summary>Gets the stable operation-row identifier.</summary>
    public string Id { get; }
    /// <summary>Gets the package that owns the behavior.</summary>
    public string PackageId { get; }
    /// <summary>Gets the source or primary format identifier.</summary>
    public string FormatId { get; }
    /// <summary>Gets the target format for directional operations, or an empty string.</summary>
    public string TargetFormatId { get; }
    /// <summary>Gets the stable identifier from the detailed owning capability contract.</summary>
    public string CapabilityId { get; }
    /// <summary>Gets the package-neutral operation.</summary>
    public OfficeOperationKind Operation { get; }
    /// <summary>Gets the supported outcome.</summary>
    public OfficeOperationSupportState State { get; }
    /// <summary>Gets the public API or catalog entry point.</summary>
    public string PublicApi { get; }
    /// <summary>Gets the named evidence supporting this row.</summary>
    public string Evidence { get; }
    /// <summary>Gets the detailed catalog that owns the claim.</summary>
    public string SourceCatalog { get; }
    /// <summary>Gets filename extensions associated with the source or primary format.</summary>
    public IReadOnlyList<string> Extensions { get; }
    /// <summary>Gets the known boundary or limitation.</summary>
    public string Limitation { get; }

    private static string NormalizeExtension(string value) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Extensions cannot contain empty values.", nameof(value));
        string normalized = value.Trim();
        return normalized[0] == '.' ? normalized.ToLowerInvariant() : "." + normalized.ToLowerInvariant();
    }
}

/// <summary>A deterministic package-neutral operation catalog.</summary>
public sealed class OfficeOperationCatalog {
    private readonly IReadOnlyDictionary<string, OfficeOperationCapability> _byId;

    /// <summary>Creates a versioned package-neutral operation catalog.</summary>
    public OfficeOperationCatalog(string id, int schemaVersion, IEnumerable<OfficeOperationCapability> capabilities) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Catalog id cannot be empty.", nameof(id));
        if (schemaVersion <= 0) throw new ArgumentOutOfRangeException(nameof(schemaVersion));
        if (capabilities == null) throw new ArgumentNullException(nameof(capabilities));

        OfficeOperationCapability[] rows = capabilities
            .OrderBy(static row => row.PackageId, StringComparer.Ordinal)
            .ThenBy(static row => row.FormatId, StringComparer.Ordinal)
            .ThenBy(static row => row.Operation)
            .ThenBy(static row => row.Id, StringComparer.Ordinal)
            .ToArray();
        if (rows.Length == 0) throw new ArgumentException("An operation catalog must contain at least one row.", nameof(capabilities));
        string[] duplicateIds = rows.GroupBy(static row => row.Id, StringComparer.Ordinal)
            .Where(static group => group.Count() > 1)
            .Select(static group => group.Key)
            .ToArray();
        if (duplicateIds.Length != 0) {
            throw new ArgumentException("Operation ids must be unique: " + string.Join(", ", duplicateIds), nameof(capabilities));
        }

        Id = id.Trim();
        SchemaVersion = schemaVersion;
        Capabilities = new ReadOnlyCollection<OfficeOperationCapability>(rows);
        _byId = new ReadOnlyDictionary<string, OfficeOperationCapability>(
            rows.ToDictionary(static row => row.Id, StringComparer.Ordinal));
    }

    /// <summary>Gets the stable catalog identifier.</summary>
    public string Id { get; }
    /// <summary>Gets the catalog schema version.</summary>
    public int SchemaVersion { get; }
    /// <summary>Gets rows in deterministic package, format, operation, and identifier order.</summary>
    public IReadOnlyList<OfficeOperationCapability> Capabilities { get; }

    /// <summary>Gets one row by exact stable identifier.</summary>
    public OfficeOperationCapability Get(string id) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Operation id cannot be empty.", nameof(id));
        if (!_byId.TryGetValue(id.Trim(), out OfficeOperationCapability? row)) {
            throw new KeyNotFoundException($"Operation '{id}' is not present in catalog '{Id}'.");
        }
        return row;
    }

    /// <summary>Gets rows owned by one package.</summary>
    public IReadOnlyList<OfficeOperationCapability> FindByPackageId(string packageId) {
        if (string.IsNullOrWhiteSpace(packageId)) throw new ArgumentException("Package id cannot be empty.", nameof(packageId));
        return Capabilities.Where(row => string.Equals(row.PackageId, packageId.Trim(), StringComparison.Ordinal)).ToArray();
    }

    /// <summary>Gets rows associated with one filename extension.</summary>
    public IReadOnlyList<OfficeOperationCapability> FindByExtension(string extension) {
        string normalized = string.IsNullOrWhiteSpace(extension)
            ? throw new ArgumentException("Extension cannot be empty.", nameof(extension))
            : extension.Trim();
        if (normalized[0] != '.') normalized = "." + normalized;
        return Capabilities.Where(row => row.Extensions.Contains(normalized, StringComparer.OrdinalIgnoreCase)).ToArray();
    }

    /// <summary>Serializes the catalog as deterministic JSON without a JSON runtime dependency.</summary>
    public string ToJson() {
        var output = new StringBuilder();
        output.Append("{\n  \"id\":\"").Append(EscapeJson(Id)).Append("\",\n  \"schemaVersion\":")
            .Append(SchemaVersion.ToString(CultureInfo.InvariantCulture)).Append(",\n  \"capabilities\":[\n");
        for (int index = 0; index < Capabilities.Count; index++) {
            OfficeOperationCapability row = Capabilities[index];
            output.Append("    {\n")
                .Append("      \"id\":\"").Append(EscapeJson(row.Id)).Append("\",\n")
                .Append("      \"packageId\":\"").Append(EscapeJson(row.PackageId)).Append("\",\n")
                .Append("      \"formatId\":\"").Append(EscapeJson(row.FormatId)).Append("\",\n")
                .Append("      \"targetFormatId\":\"").Append(EscapeJson(row.TargetFormatId)).Append("\",\n")
                .Append("      \"capabilityId\":\"").Append(EscapeJson(row.CapabilityId)).Append("\",\n")
                .Append("      \"operation\":\"").Append(row.Operation).Append("\",\n")
                .Append("      \"state\":\"").Append(row.State).Append("\",\n")
                .Append("      \"publicApi\":\"").Append(EscapeJson(row.PublicApi)).Append("\",\n")
                .Append("      \"evidence\":\"").Append(EscapeJson(row.Evidence)).Append("\",\n")
                .Append("      \"sourceCatalog\":\"").Append(EscapeJson(row.SourceCatalog)).Append("\",\n")
                .Append("      \"extensions\":[");
            for (int extensionIndex = 0; extensionIndex < row.Extensions.Count; extensionIndex++) {
                if (extensionIndex != 0) output.Append(',');
                output.Append('"').Append(EscapeJson(row.Extensions[extensionIndex])).Append('"');
            }
            output.Append("],\n      \"limitation\":\"").Append(EscapeJson(row.Limitation)).Append("\"\n    }");
            if (index + 1 < Capabilities.Count) output.Append(',');
            output.Append('\n');
        }
        return output.Append("  ]\n}").ToString();
    }

    /// <summary>Formats the catalog as a deterministic Markdown support matrix.</summary>
    public string ToMarkdown() {
        var output = new StringBuilder();
        output.Append("# ").Append(Id).Append(" operation contract\n\nSchema version: ")
            .Append(SchemaVersion.ToString(CultureInfo.InvariantCulture))
            .Append("\n\nEach row retains its detailed owning catalog and evidence. A supported row may still carry a bounded limitation.\n\n")
            .Append("| Package | Format | Target | Operation | State | Capability | Source contract | Public API | Evidence | Boundary |\n")
            .Append("| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |\n");
        foreach (OfficeOperationCapability row in Capabilities) {
            output.Append("| `").Append(EscapeMarkdown(row.PackageId)).Append("` | ")
                .Append(EscapeMarkdown(row.FormatId)).Append(" | ")
                .Append(EscapeMarkdown(row.TargetFormatId)).Append(" | ")
                .Append(row.Operation).Append(" | ").Append(row.State).Append(" | ")
                .Append(EscapeMarkdown(row.CapabilityId)).Append(" | ")
                .Append(EscapeMarkdown(row.SourceCatalog)).Append(" | `")
                .Append(EscapeMarkdown(row.PublicApi)).Append("` | ")
                .Append(EscapeMarkdown(row.Evidence)).Append(" | ")
                .Append(EscapeMarkdown(row.Limitation)).Append(" |\n");
        }
        return output.ToString();
    }

    private static string EscapeJson(string value) {
        var escaped = new StringBuilder(value.Length + 8);
        foreach (char character in value) {
            switch (character) {
                case '"': escaped.Append("\\\""); break;
                case '\\': escaped.Append("\\\\"); break;
                case '\b': escaped.Append("\\b"); break;
                case '\f': escaped.Append("\\f"); break;
                case '\n': escaped.Append("\\n"); break;
                case '\r': escaped.Append("\\r"); break;
                case '\t': escaped.Append("\\t"); break;
                default:
                    if (character < ' ') escaped.Append("\\u").Append(((int)character).ToString("x4", CultureInfo.InvariantCulture));
                    else escaped.Append(character);
                    break;
            }
        }
        return escaped.ToString();
    }

    private static string EscapeMarkdown(string value) => value
        .Replace("\\", "\\\\")
        .Replace("|", "\\|")
        .Replace("\r", " ")
        .Replace("\n", " ");
}
