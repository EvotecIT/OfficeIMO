using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>One versioned, machine-readable feature row in an Office legacy/modern compatibility contract.</summary>
public sealed class OfficeCapability {
    /// <summary>Creates a compatibility capability row.</summary>
    public OfficeCapability(
        string id,
        string formatId,
        OfficeDocumentFamily family,
        string category,
        string description,
        OfficeCapabilityRepresentability representability,
        OfficeCapabilityCoverageState legacyImport,
        OfficeCapabilityCoverageState newLegacyWrite,
        OfficeCapabilityCoverageState legacyRoundTrip,
        OfficeCapabilityCoverageState modernToLegacy,
        OfficeCapabilityCoverageState legacyToModern,
        OfficeCompatibilityImpact affectedFidelity = OfficeCompatibilityImpact.None,
        string? note = null) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Capability id cannot be empty.", nameof(id));
        if (string.IsNullOrWhiteSpace(formatId)) throw new ArgumentException("Format id cannot be empty.", nameof(formatId));
        if (string.IsNullOrWhiteSpace(category)) throw new ArgumentException("Capability category cannot be empty.", nameof(category));
        if (string.IsNullOrWhiteSpace(description)) throw new ArgumentException("Capability description cannot be empty.", nameof(description));

        Id = id.Trim();
        FormatId = formatId.Trim();
        Family = family;
        Category = category.Trim();
        Description = description.Trim();
        Representability = representability;
        LegacyImport = legacyImport;
        NewLegacyWrite = newLegacyWrite;
        LegacyRoundTrip = legacyRoundTrip;
        ModernToLegacy = modernToLegacy;
        LegacyToModern = legacyToModern;
        AffectedFidelity = affectedFidelity;
        Note = note?.Trim() ?? string.Empty;
    }

    /// <summary>Gets the stable capability identifier.</summary>
    public string Id { get; }

    /// <summary>Gets the stable legacy format identifier this row describes.</summary>
    public string FormatId { get; }

    /// <summary>Gets the owning document family.</summary>
    public OfficeDocumentFamily Family { get; }

    /// <summary>Gets the report category.</summary>
    public string Category { get; }

    /// <summary>Gets the user-facing feature description.</summary>
    public string Description { get; }

    /// <summary>Gets what the legacy format itself can represent.</summary>
    public OfficeCapabilityRepresentability Representability { get; }

    /// <summary>Gets current legacy import coverage.</summary>
    public OfficeCapabilityCoverageState LegacyImport { get; }

    /// <summary>Gets current new legacy authoring coverage.</summary>
    public OfficeCapabilityCoverageState NewLegacyWrite { get; }

    /// <summary>Gets current legacy round-trip coverage.</summary>
    public OfficeCapabilityCoverageState LegacyRoundTrip { get; }

    /// <summary>Gets current modern-to-legacy conversion coverage.</summary>
    public OfficeCapabilityCoverageState ModernToLegacy { get; }

    /// <summary>Gets current legacy-to-modern conversion coverage.</summary>
    public OfficeCapabilityCoverageState LegacyToModern { get; }

    /// <summary>Gets the fidelity dimensions affected when a non-native mapping is used.</summary>
    public OfficeCompatibilityImpact AffectedFidelity { get; }

    /// <summary>Gets important limitations or fallback behavior.</summary>
    public string Note { get; }

    /// <summary>Gets whether at least one lane is not implemented.</summary>
    public bool HasUnimplementedCoverage => GetStates().Contains(OfficeCapabilityCoverageState.NotImplemented);

    /// <summary>Gets whether at least one lane deliberately blocks the operation.</summary>
    public bool HasBlockedCoverage => GetStates().Contains(OfficeCapabilityCoverageState.Blocked);

    /// <summary>Gets whether all lanes use native or equivalent editable representations.</summary>
    public bool IsFullyEditableAcrossLanes => GetStates().All(static state =>
        state == OfficeCapabilityCoverageState.Native || state == OfficeCapabilityCoverageState.Equivalent);

    /// <summary>Gets the coverage state for a direction.</summary>
    public OfficeCapabilityCoverageState GetState(OfficeCapabilityLane lane) => lane switch {
        OfficeCapabilityLane.LegacyImport => LegacyImport,
        OfficeCapabilityLane.NewLegacyWrite => NewLegacyWrite,
        OfficeCapabilityLane.LegacyRoundTrip => LegacyRoundTrip,
        OfficeCapabilityLane.ModernToLegacy => ModernToLegacy,
        OfficeCapabilityLane.LegacyToModern => LegacyToModern,
        _ => throw new ArgumentOutOfRangeException(nameof(lane))
    };

    private IEnumerable<OfficeCapabilityCoverageState> GetStates() {
        yield return LegacyImport;
        yield return NewLegacyWrite;
        yield return LegacyRoundTrip;
        yield return ModernToLegacy;
        yield return LegacyToModern;
    }
}

/// <summary>A deterministic capability contract for one Office binary format or format family.</summary>
public sealed class OfficeCapabilityCatalog {
    private readonly IReadOnlyDictionary<string, OfficeCapability> _capabilitiesById;

    /// <summary>Creates a versioned capability catalog.</summary>
    public OfficeCapabilityCatalog(string id, int schemaVersion, IEnumerable<OfficeCapability> capabilities) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Catalog id cannot be empty.", nameof(id));
        if (schemaVersion <= 0) throw new ArgumentOutOfRangeException(nameof(schemaVersion));
        if (capabilities == null) throw new ArgumentNullException(nameof(capabilities));

        OfficeCapability[] rows = capabilities.ToArray();
        if (rows.Length == 0) throw new ArgumentException("A capability catalog must contain at least one row.", nameof(capabilities));
        string[] duplicateIds = rows.GroupBy(row => row.Id, StringComparer.Ordinal)
            .Where(group => group.Count() > 1)
            .Select(group => group.Key)
            .ToArray();
        if (duplicateIds.Length > 0) {
            throw new ArgumentException($"Capability ids must be unique. Duplicate ids: {string.Join(", ", duplicateIds)}.", nameof(capabilities));
        }

        Id = id.Trim();
        SchemaVersion = schemaVersion;
        Capabilities = new ReadOnlyCollection<OfficeCapability>(rows);
        _capabilitiesById = new ReadOnlyDictionary<string, OfficeCapability>(
            rows.ToDictionary(row => row.Id, StringComparer.Ordinal));
    }

    /// <summary>Gets the stable catalog identifier.</summary>
    public string Id { get; }

    /// <summary>Gets the contract schema version.</summary>
    public int SchemaVersion { get; }

    /// <summary>Gets capability rows in stable order.</summary>
    public IReadOnlyList<OfficeCapability> Capabilities { get; }

    /// <summary>Gets whether any lane still lacks an implementation.</summary>
    public bool HasUnimplementedCoverage => Capabilities.Any(static capability => capability.HasUnimplementedCoverage);

    /// <summary>Gets whether any operation is deliberately blocked for format-safety reasons.</summary>
    public bool HasBlockedCoverage => Capabilities.Any(static capability => capability.HasBlockedCoverage);

    /// <summary>Gets rows that still contain an unimplemented lane.</summary>
    public IReadOnlyList<OfficeCapability> UnimplementedCapabilities => Capabilities
        .Where(static capability => capability.HasUnimplementedCoverage)
        .ToArray();

    /// <summary>Gets a capability by its exact stable id.</summary>
    public OfficeCapability Get(string id) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Capability id cannot be empty.", nameof(id));
        if (!_capabilitiesById.TryGetValue(id.Trim(), out OfficeCapability? capability)) {
            throw new KeyNotFoundException($"Capability '{id}' is not present in catalog '{Id}'.");
        }
        return capability;
    }

    /// <summary>Serializes the complete capability contract as deterministic JSON.</summary>
    public string ToJson(bool indented = true) {
        string newline = indented ? "\n" : string.Empty;
        string i1 = indented ? "  " : string.Empty;
        string i2 = indented ? "    " : string.Empty;
        string i3 = indented ? "      " : string.Empty;
        var json = new StringBuilder();
        json.Append('{').Append(newline)
            .Append(i1).Append("\"id\":\"").Append(EscapeJson(Id)).Append("\",").Append(newline)
            .Append(i1).Append("\"schemaVersion\":").Append(SchemaVersion).Append(',').Append(newline)
            .Append(i1).Append("\"hasUnimplementedCoverage\":")
            .Append(HasUnimplementedCoverage ? "true" : "false").Append(',').Append(newline)
            .Append(i1).Append("\"capabilities\": [").Append(newline);

        for (int index = 0; index < Capabilities.Count; index++) {
            OfficeCapability row = Capabilities[index];
            json.Append(i2).Append('{').Append(newline)
                .Append(i3).Append("\"id\":\"").Append(EscapeJson(row.Id)).Append("\",").Append(newline)
                .Append(i3).Append("\"formatId\":\"").Append(EscapeJson(row.FormatId)).Append("\",").Append(newline)
                .Append(i3).Append("\"family\":\"").Append(row.Family).Append("\",").Append(newline)
                .Append(i3).Append("\"category\":\"").Append(EscapeJson(row.Category)).Append("\",").Append(newline)
                .Append(i3).Append("\"description\":\"").Append(EscapeJson(row.Description)).Append("\",").Append(newline)
                .Append(i3).Append("\"representability\":\"").Append(row.Representability).Append("\",").Append(newline)
                .Append(i3).Append("\"legacyImport\":\"").Append(row.LegacyImport).Append("\",").Append(newline)
                .Append(i3).Append("\"newLegacyWrite\":\"").Append(row.NewLegacyWrite).Append("\",").Append(newline)
                .Append(i3).Append("\"legacyRoundTrip\":\"").Append(row.LegacyRoundTrip).Append("\",").Append(newline)
                .Append(i3).Append("\"modernToLegacy\":\"").Append(row.ModernToLegacy).Append("\",").Append(newline)
                .Append(i3).Append("\"legacyToModern\":\"").Append(row.LegacyToModern).Append("\",").Append(newline)
                .Append(i3).Append("\"affectedFidelity\":\"").Append(row.AffectedFidelity).Append("\",").Append(newline)
                .Append(i3).Append("\"note\":\"").Append(EscapeJson(row.Note)).Append('"').Append(newline)
                .Append(i2).Append('}');
            if (index + 1 < Capabilities.Count) json.Append(',');
            json.Append(newline);
        }

        json.Append(i1).Append(']').Append(newline).Append('}');
        return json.ToString();
    }

    /// <summary>Formats the capability contract as a deterministic Markdown table.</summary>
    public string ToMarkdown() {
        var markdown = new StringBuilder();
        markdown.Append("# ").Append(Id).AppendLine(" capability contract");
        markdown.AppendLine();
        markdown.Append("Schema version: ").Append(SchemaVersion).AppendLine();
        markdown.AppendLine();
        markdown.AppendLine("| Category | Capability | Format | Representation | Legacy import | New legacy | Legacy round-trip | Modern to legacy | Legacy to modern | Fidelity | Note |");
        markdown.AppendLine("| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |");
        foreach (OfficeCapability row in Capabilities) {
            markdown.Append("| ").Append(EscapeMarkdown(row.Category))
                .Append(" | ").Append(EscapeMarkdown(row.Id))
                .Append(" | ").Append(EscapeMarkdown(row.FormatId))
                .Append(" | ").Append(row.Representability)
                .Append(" | ").Append(row.LegacyImport)
                .Append(" | ").Append(row.NewLegacyWrite)
                .Append(" | ").Append(row.LegacyRoundTrip)
                .Append(" | ").Append(row.ModernToLegacy)
                .Append(" | ").Append(row.LegacyToModern)
                .Append(" | ").Append(row.AffectedFidelity)
                .Append(" | ").Append(EscapeMarkdown(row.Note)).AppendLine(" |");
        }
        return NormalizeNewlines(markdown.ToString());
    }

    private static string EscapeJson(string value) => (value ?? string.Empty)
        .Replace("\\", "\\\\").Replace("\"", "\\\"")
        .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\t", "\\t");

    private static string EscapeMarkdown(string value) => (value ?? string.Empty)
        .Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");

    private static string NormalizeNewlines(string value) => value.Replace("\r\n", "\n").Replace("\r", "\n");
}
