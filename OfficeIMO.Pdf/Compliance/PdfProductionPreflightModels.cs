using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Inspection policy for a production PDF; these profiles do not certify standards conformance.</summary>
public enum PdfProductionPreflightProfile {
    /// <summary>General print preparation with warnings for missing production metadata.</summary>
    GeneralPrint,
    /// <summary>Conservative candidate checks before independent PDF/X-1a validation.</summary>
    PdfX1aCandidate,
    /// <summary>Conservative candidate checks before independent PDF/X-4 validation.</summary>
    PdfX4Candidate
}

/// <summary>Severity of one production finding.</summary>
public enum PdfProductionFindingSeverity {
    /// <summary>Review is recommended under this profile.</summary>
    Warning,
    /// <summary>The artifact fails a profile rule.</summary>
    Error,
    /// <summary>Evidence was incomplete; the condition remains unknown.</summary>
    Indeterminate
}

/// <summary>Machine-readable production finding.</summary>
public enum PdfProductionFindingKind {
    /// <summary>No catalog output intent is present.</summary>
    MissingOutputIntent,
    /// <summary>The output intent is incomplete or incompatible with the profile.</summary>
    InvalidOutputIntent,
    /// <summary>Production boundary boxes are missing or inconsistent.</summary>
    InvalidPageBoxes,
    /// <summary>A reachable font lacks a self-contained program.</summary>
    UnembeddedFont,
    /// <summary>A font selection context could not be inspected.</summary>
    UninspectableFont,
    /// <summary>Reachable page content uses device RGB.</summary>
    DeviceRgbColor,
    /// <summary>Reachable page content uses device-independent color forbidden by the selected candidate profile.</summary>
    DeviceIndependentColor,
    /// <summary>Reachable page content uses transparency.</summary>
    Transparency,
    /// <summary>Color content could not be inspected completely.</summary>
    UninspectableColor,
    /// <summary>A placed image is below the selected effective resolution.</summary>
    LowImageResolution,
    /// <summary>A placed image lacks complete resolution evidence.</summary>
    UninspectableImageResolution
}

/// <summary>Bounded settings for production inspection.</summary>
public sealed class PdfProductionPreflightOptions {
    /// <summary>Rule profile; independent validation remains necessary for standards claims.</summary>
    public PdfProductionPreflightProfile Profile { get; set; } = PdfProductionPreflightProfile.GeneralPrint;
    /// <summary>Pages to inspect; null selects the complete document.</summary>
    public PdfPageSelection? PageSelection { get; set; }
    /// <summary>Minimum effective image pixels per inch; null uses the profile default.</summary>
    public double? MinimumImagePpi { get; set; }
    /// <summary>Maximum selected pages.</summary>
    public int MaxPages { get; set; } = 100;
    /// <summary>Maximum retained findings.</summary>
    public int MaxFindings { get; set; } = 1_000;
    /// <summary>Maximum proposed page-box metadata fixups.</summary>
    public int MaxFixupProposals { get; set; } = 200;

    internal double EffectiveMinimumImagePpi => MinimumImagePpi ??
        (Profile == PdfProductionPreflightProfile.GeneralPrint ? 200D : 300D);

    internal void Validate() {
        if (Profile is not (PdfProductionPreflightProfile.GeneralPrint or PdfProductionPreflightProfile.PdfX1aCandidate or PdfProductionPreflightProfile.PdfX4Candidate)) {
            throw new ArgumentOutOfRangeException(nameof(Profile));
        }
        if (MinimumImagePpi.HasValue && (double.IsNaN(MinimumImagePpi.Value) || double.IsInfinity(MinimumImagePpi.Value) || MinimumImagePpi.Value <= 0D)) {
            throw new ArgumentOutOfRangeException(nameof(MinimumImagePpi));
        }
        if (MaxPages <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPages));
        if (MaxFindings <= 0) throw new ArgumentOutOfRangeException(nameof(MaxFindings));
        if (MaxFixupProposals < 0) throw new ArgumentOutOfRangeException(nameof(MaxFixupProposals));
    }

    internal PdfProductionPreflightOptions Copy() => new PdfProductionPreflightOptions {
        Profile = Profile, PageSelection = PageSelection, MinimumImagePpi = MinimumImagePpi,
        MaxPages = MaxPages, MaxFindings = MaxFindings, MaxFixupProposals = MaxFixupProposals
    };
}

/// <summary>One document-wide or page-linked production concern.</summary>
public sealed class PdfProductionFinding {
    internal PdfProductionFinding(PdfProductionFindingKind kind, PdfProductionFindingSeverity severity, int? pageNumber,
        string message, PdfLogicalVisualBounds? visualBounds = null, double? observedImagePpi = null) {
        Kind = kind; Severity = severity; PageNumber = pageNumber; Message = message;
        VisualBounds = visualBounds; ObservedImagePpi = observedImagePpi;
    }
    /// <summary>Stable finding category.</summary>
    public PdfProductionFindingKind Kind { get; }
    /// <summary>Severity under the selected profile.</summary>
    public PdfProductionFindingSeverity Severity { get; }
    /// <summary>One-based source page, or null for document-level evidence.</summary>
    public int? PageNumber { get; }
    /// <summary>Review detail.</summary>
    public string Message { get; }
    /// <summary>Top-left visual bounds for a placed image, when available.</summary>
    public PdfLogicalVisualBounds? VisualBounds { get; }
    /// <summary>Minimum effective pixels per inch for a low-resolution image placement.</summary>
    public double? ObservedImagePpi { get; }
}

/// <summary>Reviewable metadata-only page-box change; it never extends page artwork.</summary>
public sealed class PdfProductionFixupProposal {
    internal PdfProductionFixupProposal(int index, int pageNumber, PdfPageBoundaryBox box, PdfPageBox bounds, string reason) {
        Index = index; PageNumber = pageNumber; Box = box; Bounds = bounds; Reason = reason;
    }
    /// <summary>Stable index for explicit selection.</summary>
    public int Index { get; }
    /// <summary>One-based source page.</summary>
    public int PageNumber { get; }
    /// <summary>TrimBox or BleedBox to set.</summary>
    public PdfPageBoundaryBox Box { get; }
    /// <summary>Proposed PDF user-space rectangle.</summary>
    public PdfPageBox Bounds { get; }
    /// <summary>Why the metadata change was proposed.</summary>
    public string Reason { get; }
}

/// <summary>Exact-artifact production inspection and separately selectable metadata fixups.</summary>
public sealed class PdfProductionPreflightReport {
    private readonly byte[] _analyzedPdf;
    private readonly PdfLoadOptions _readOptions;
    private readonly PdfProductionPreflightOptions _options;

    internal PdfProductionPreflightReport(byte[] analyzedPdf, PdfLoadOptions readOptions, PdfProductionPreflightOptions options,
        IReadOnlyList<int> inspectedPages, IReadOnlyList<PdfProductionFinding> findings,
        IReadOnlyList<PdfProductionFixupProposal> fixups) {
        _analyzedPdf = (byte[])analyzedPdf.Clone();
        _readOptions = readOptions;
        _options = options.Copy();
        SourceSha256 = PdfArtifactFingerprint.ComputeSha256(_analyzedPdf);
        Profile = options.Profile;
        InspectedPages = Array.AsReadOnly(inspectedPages.ToArray());
        Findings = Array.AsReadOnly(findings.ToArray());
        FixupProposals = Array.AsReadOnly(fixups.ToArray());
    }
    /// <summary>Profile used for this report.</summary>
    public PdfProductionPreflightProfile Profile { get; }
    /// <summary>SHA-256 of the exact inspected PDF.</summary>
    public string SourceSha256 { get; }
    /// <summary>One-based pages inspected.</summary>
    public IReadOnlyList<int> InspectedPages { get; }
    /// <summary>Findings, including explicit incomplete evidence.</summary>
    public IReadOnlyList<PdfProductionFinding> Findings { get; }
    /// <summary>Optional page-box metadata proposals. No change occurs until selected.</summary>
    public IReadOnlyList<PdfProductionFixupProposal> FixupProposals { get; }
    /// <summary>Whether this inspection found any error-level concerns. This does not certify conformance.</summary>
    public bool HasErrors => Findings.Any(static finding => finding.Severity == PdfProductionFindingSeverity.Error);

    /// <summary>Applies selected metadata proposals to the analyzed snapshot and reopens the result for inspection.</summary>
    public PdfProductionFixupResult ApplySelected(IReadOnlyCollection<int> proposalIndices, CancellationToken cancellationToken = default) {
        Guard.NotNull(proposalIndices, nameof(proposalIndices));
        cancellationToken.ThrowIfCancellationRequested();
        if (proposalIndices.Count == 0) throw new ArgumentException("Select at least one proposal.", nameof(proposalIndices));
        int[] selected = proposalIndices.Distinct().OrderBy(static index => index).ToArray();
        if (selected.Length != proposalIndices.Count || selected.Any(index => index < 0 || index >= FixupProposals.Count)) {
            throw new ArgumentOutOfRangeException(nameof(proposalIndices), "Indices must be distinct and present in this report.");
        }
        PdfProductionFixupProposal[] proposals = selected.Select(index => FixupProposals[index]).ToArray();
        byte[] rewritten = PdfPageEditor.SetPageBoxesWithReadOptions(_analyzedPdf, proposals, _readOptions, cancellationToken);
        PdfLoadOptions outputReadOptions = PdfLoadOptions.ForGeneratedOutput(_readOptions, _analyzedPdf, rewritten);
        PdfDocument output = PdfDocument.Load(rewritten, outputReadOptions);
        PdfProductionPreflightReport after = PdfProductionPreflightInspector.Inspect(output, _options, cancellationToken);
        return new PdfProductionFixupResult(output, after);
    }
}

/// <summary>Applied PDF plus an independently reopened engine inspection of its exact output.</summary>
public sealed class PdfProductionFixupResult {
    internal PdfProductionFixupResult(PdfDocument document, PdfProductionPreflightReport after) {
        Document = document; After = after;
    }
    /// <summary>Rewritten document.</summary>
    public PdfDocument Document { get; }
    /// <summary>Inspection of the rewritten artifact; external standards validation remains separate.</summary>
    public PdfProductionPreflightReport After { get; }
}
