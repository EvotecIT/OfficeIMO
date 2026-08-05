using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>Immutable feature-level fidelity report for one Office conversion assessment.</summary>
public sealed class OfficeCompatibilityReport {
    private readonly ReadOnlyCollection<OfficeCompatibilityFinding> _findings;

    /// <summary>Creates a compatibility report.</summary>
    public OfficeCompatibilityReport(
        OfficeFormatDescriptor sourceFormat,
        OfficeFormatDescriptor destinationFormat,
        OfficeCompatibilityMode mode,
        IEnumerable<OfficeCompatibilityFinding>? findings = null) {
        SourceFormat = sourceFormat ?? throw new ArgumentNullException(nameof(sourceFormat));
        DestinationFormat = destinationFormat ?? throw new ArgumentNullException(nameof(destinationFormat));
        if (sourceFormat.Family != destinationFormat.Family) {
            throw new ArgumentException("Source and destination formats must belong to the same Office document family.", nameof(destinationFormat));
        }

        Mode = mode;
        _findings = Array.AsReadOnly((findings ?? Array.Empty<OfficeCompatibilityFinding>()).ToArray());
    }

    /// <summary>Gets the detected source format.</summary>
    public OfficeFormatDescriptor SourceFormat { get; }

    /// <summary>Gets the requested destination format.</summary>
    public OfficeFormatDescriptor DestinationFormat { get; }

    /// <summary>Gets the requested compatibility policy.</summary>
    public OfficeCompatibilityMode Mode { get; }

    /// <summary>Gets all feature-level decisions in discovery order.</summary>
    public IReadOnlyList<OfficeCompatibilityFinding> Findings => _findings;

    /// <summary>Gets whether any finding reports fidelity loss.</summary>
    public bool HasLoss => _findings.Any(finding => finding.RepresentsLoss);

    /// <summary>Gets whether any feature blocks artifact creation.</summary>
    public bool HasBlockedFeatures => _findings.Any(finding => finding.State == OfficeCompatibilityState.Blocked);

    /// <summary>Gets whether active content or another security property changes.</summary>
    public bool HasSecurityImpact => _findings.Any(finding =>
        (finding.Impact & OfficeCompatibilityImpact.Security) != 0);

    /// <summary>Gets whether every reported feature remains native or semantically equivalent.</summary>
    public bool IsStrictlyCompatible => !_findings.Any(finding =>
        finding.State != OfficeCompatibilityState.Native
        && finding.State != OfficeCompatibilityState.Equivalent);

    /// <summary>Returns findings affecting a requested fidelity dimension.</summary>
    public IReadOnlyList<OfficeCompatibilityFinding> GetFindings(OfficeCompatibilityImpact impact) {
        if (impact == OfficeCompatibilityImpact.None) {
            return Array.AsReadOnly(_findings.Where(finding => finding.Impact == OfficeCompatibilityImpact.None).ToArray());
        }

        return Array.AsReadOnly(_findings.Where(finding => (finding.Impact & impact) != 0).ToArray());
    }

    /// <summary>Throws when the report contains loss or a blocked feature.</summary>
    public void RequireNoLoss() {
        if (HasBlockedFeatures) throw new InvalidOperationException("Office conversion contains one or more blocked features.");
        if (HasLoss) throw new InvalidOperationException("Office conversion contains one or more lossy feature mappings.");
    }
}
