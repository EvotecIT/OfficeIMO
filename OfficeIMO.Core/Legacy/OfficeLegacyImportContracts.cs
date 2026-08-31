using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO;

/// <summary>Describes whether a legacy import reconstructed source structure or recovered only safe salvage content.</summary>
public enum OfficeLegacyImportQuality {
    /// <summary>The selected adapter recovered a structured model for its documented source profile.</summary>
    Structured,

    /// <summary>The selected adapter recovered bounded text or tabular content without claiming complete source structure.</summary>
    Salvage
}

/// <summary>Identifies active or externally resolved source content that an importer kept inert.</summary>
[Flags]
public enum OfficeLegacyInertContentKind {
    /// <summary>No active or externally resolved content was discovered.</summary>
    None = 0,
    /// <summary>Source macro content was discovered but never executed or projected as executable code.</summary>
    Macros = 1,
    /// <summary>Embedded scripts, programs, or executable code were discovered but never executed.</summary>
    EmbeddedCode = 2,
    /// <summary>External links or data connections were discovered but never resolved.</summary>
    ExternalLinks = 4,
    /// <summary>Embedded OLE or application objects were discovered but never activated.</summary>
    EmbeddedObjects = 8
}

/// <summary>Hard resource limits shared by managed legacy importers.</summary>
public sealed class OfficeLegacyImportLimits {
    /// <summary>Gets or sets the maximum source size in bytes.</summary>
    public int MaxInputBytes { get; set; } = 64 * 1024 * 1024;

    /// <summary>Gets or sets the maximum number of recovered text characters.</summary>
    public int MaxTextCharacters { get; set; } = 4 * 1024 * 1024;

    /// <summary>Gets or sets the maximum number of logical records inspected.</summary>
    public int MaxRecords { get; set; } = 1_000_000;

    /// <summary>Gets or sets the maximum number of recovered document blocks or spreadsheet cells.</summary>
    public int MaxItems { get; set; } = 250_000;

    /// <summary>Gets or sets the maximum number of compound-document streams inspected.</summary>
    public int MaxCompoundStreams { get; set; } = 512;

    /// <summary>Creates an independent copy of these limits.</summary>
    public OfficeLegacyImportLimits Clone() => new() {
        MaxInputBytes = MaxInputBytes,
        MaxTextCharacters = MaxTextCharacters,
        MaxRecords = MaxRecords,
        MaxItems = MaxItems,
        MaxCompoundStreams = MaxCompoundStreams
    };

    /// <summary>Throws when a configured limit is outside the supported positive range.</summary>
    public void Validate() {
        if (MaxInputBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxInputBytes));
        if (MaxTextCharacters < 1) throw new ArgumentOutOfRangeException(nameof(MaxTextCharacters));
        if (MaxRecords < 1) throw new ArgumentOutOfRangeException(nameof(MaxRecords));
        if (MaxItems < 1) throw new ArgumentOutOfRangeException(nameof(MaxItems));
        if (MaxCompoundStreams < 1) throw new ArgumentOutOfRangeException(nameof(MaxCompoundStreams));
    }
}

/// <summary>Immutable import and loss report produced by a managed legacy source adapter.</summary>
public sealed class OfficeLegacyImportReport {
    private readonly ReadOnlyCollection<OfficeCompatibilityFinding> _findings;

    /// <summary>Creates a legacy import report.</summary>
    public OfficeLegacyImportReport(
        string sourceFormatId,
        OfficeLegacyImportQuality quality,
        IEnumerable<OfficeCompatibilityFinding>? findings = null,
        OfficeLegacyInertContentKind inertContent = OfficeLegacyInertContentKind.None,
        int recoveredItemCount = 0) {
        if (string.IsNullOrWhiteSpace(sourceFormatId)) {
            throw new ArgumentException("Source format id cannot be empty.", nameof(sourceFormatId));
        }
        if (recoveredItemCount < 0) throw new ArgumentOutOfRangeException(nameof(recoveredItemCount));

        SourceFormatId = sourceFormatId.Trim();
        Quality = quality;
        InertContent = inertContent;
        RecoveredItemCount = recoveredItemCount;
        _findings = Array.AsReadOnly((findings ?? Array.Empty<OfficeCompatibilityFinding>()).ToArray());
    }

    /// <summary>Gets the stable identifier of the detected legacy source profile.</summary>
    public string SourceFormatId { get; }

    /// <summary>Gets the achieved import quality.</summary>
    public OfficeLegacyImportQuality Quality { get; }

    /// <summary>Gets source content that was deliberately kept inert.</summary>
    public OfficeLegacyInertContentKind InertContent { get; }

    /// <summary>Gets the number of recovered blocks or cells.</summary>
    public int RecoveredItemCount { get; }

    /// <summary>Gets feature-level recovery, approximation, omission, and safety findings.</summary>
    public IReadOnlyList<OfficeCompatibilityFinding> Findings => _findings;

    /// <summary>Gets whether any finding reports source fidelity loss.</summary>
    public bool HasLoss => _findings.Any(finding => finding.RepresentsLoss);

    /// <summary>Gets whether the source contained active or externally resolved content that remained inert.</summary>
    public bool HasInertContent => InertContent != OfficeLegacyInertContentKind.None;

    /// <summary>Throws when the import contains a known loss or blocked feature.</summary>
    public void RequireStructuredNoLoss() {
        if (Quality != OfficeLegacyImportQuality.Structured) {
            throw new InvalidOperationException("Legacy import produced salvage quality rather than a structured reconstruction.");
        }
        if (_findings.Any(finding => finding.State == OfficeCompatibilityState.Blocked)) {
            throw new InvalidOperationException("Legacy import contains one or more blocked source features.");
        }
        if (HasLoss) throw new InvalidOperationException("Legacy import contains one or more lossy source mappings.");
    }
}
