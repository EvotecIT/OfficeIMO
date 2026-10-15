using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>Immutable provider evidence for one page. No correction is selected automatically.</summary>
public sealed class PdfPageOrientationProposal {
    internal PdfPageOrientationProposal(int page, int original, int? correction, double? confidence,
        bool recommended, string provider, IReadOnlyList<string> diagnostics) {
        PageNumber = page; OriginalRotationDegrees = original; ClockwiseCorrectionDegrees = correction;
        Confidence = confidence; IsRecommended = recommended; Provider = provider;
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
    }
    /// <summary>One-based page in the captured document.</summary>
    public int PageNumber { get; }
    /// <summary>Effective saved rotation before detection, including inherited rotation.</summary>
    public int OriginalRotationDegrees { get; }
    /// <summary>Clockwise quarter-turn correction, or null when evidence is unavailable or invalid.</summary>
    public int? ClockwiseCorrectionDegrees { get; }
    /// <summary>Normalized evidence from zero through one; calibration remains provider-specific.</summary>
    public double? Confidence { get; }
    /// <summary>Whether the nonzero correction meets the requested confidence threshold.</summary>
    public bool IsRecommended { get; }
    /// <summary>Whether a valid nonzero correction can be explicitly accepted by a reviewer.</summary>
    public bool CanApply => ClockwiseCorrectionDegrees is 90 or 180 or 270;
    /// <summary>Provider identity captured by the shared OCR execution boundary.</summary>
    public string Provider { get; }
    /// <summary>Bounded rendering and provider diagnostics.</summary>
    public IReadOnlyList<string> Diagnostics { get; }
}

/// <summary>A private PDF snapshot and orientation proposals that require explicit selection before mutation.</summary>
public sealed class PdfPageOrientationReview {
    private readonly PdfDocument _source;
    private readonly byte[] _sourceHash;
    private readonly PdfOcrMergeOptions _options;
    internal PdfPageOrientationReview(PdfDocument source, PdfOcrMergeOptions options,
        IReadOnlyList<PdfPageOrientationProposal> pages) {
        _source = source; _options = options;
        _sourceHash = Hash(source.GetBytesForOperation());
        Pages = Array.AsReadOnly(pages.ToArray());
    }
    /// <summary>Detection evidence in selected page order. Confidence does not imply selection.</summary>
    public IReadOnlyList<PdfPageOrientationProposal> Pages { get; }

    /// <summary>Renders the original and proposed appearances without modifying page content.</summary>
    public async Task<PdfScanPreview> PreviewAsync(PdfPageOrientationProposal proposal, CancellationToken cancellationToken = default) {
        ValidateProposal(proposal, requireCorrection: false);
        var options = _options.Clone();
        options.DetectOrientation = false; options.Regions = Array.Empty<PdfOcrPageRegion>(); options.Perspective = null;
        var processing = new OfficeScanProcessingOptions {
            ClockwiseQuarterTurns = (proposal.ClockwiseCorrectionDegrees ?? 0) / 90,
            Deskew = false, NormalizeBackground = false, ColorMode = OfficeScanColorMode.PreserveColor
        };
        processing.MaximumPixels = Math.Min(processing.MaximumPixels, options.MaxPixelsPerPage);
        options.ScanProcessing = processing;
        var preview = await _source.PreviewScanAsync(proposal.PageNumber, options, cancellationToken).ConfigureAwait(false);
        if (proposal.CanApply && preview.ScanProcessing?.Steps.Any(step => step.Operation == "orientation" && step.Applied) != true)
            throw new InvalidOperationException("The proposed orientation could not be previewed within the scan-processing budget. Reduce the review DPI. " + string.Join(" ", preview.Diagnostics));
        return preview;
    }

    /// <summary>Returns an updated document with explicitly accepted proposals applied as one atomic page-attribute batch.</summary>
    /// <remarks>The current bytes must match the reviewed snapshot. Foreign, repeated, invalid and zero corrections
    /// are rejected before mutation. Low-confidence valid evidence may be accepted after visual review.
    /// Signature and certification authorization remains owned by the PDF mutation planner.</remarks>
    public PdfOperationResult<PdfDocument> ApplyResult(PdfDocument document,
        IEnumerable<PdfPageOrientationProposal> accepted, CancellationToken cancellationToken = default) {
        Guard.NotNull(document, nameof(document)); Guard.NotNull(accepted, nameof(accepted));
        cancellationToken.ThrowIfCancellationRequested();
        if (!Hash(document.GetBytesForOperation(cancellationToken)).SequenceEqual(_sourceHash))
            throw new InvalidOperationException("The document changed after orientation detection. Review its orientation again.");
        var selected = new HashSet<PdfPageOrientationProposal>();
        foreach (var proposal in accepted) {
            cancellationToken.ThrowIfCancellationRequested(); ValidateProposal(proposal, requireCorrection: true);
            if (!selected.Add(proposal)) throw new ArgumentException("A page correction cannot be accepted twice.", nameof(accepted));
        }
        if (selected.Count == 0) throw new ArgumentException("Accept at least one page correction.", nameof(accepted));
        var edits = selected.Select(proposal => PdfPageAttributeEdit.RotateBy(proposal.ClockwiseCorrectionDegrees!.Value,
            PdfPageSelection.From(proposal.PageNumber))).ToArray();
        return document.Pages.ApplyAttributesResult(edits, cancellationToken: cancellationToken);
    }
    private void ValidateProposal(PdfPageOrientationProposal proposal, bool requireCorrection) {
        if (proposal is null || !Pages.Contains(proposal) || requireCorrection && !proposal.CanApply)
            throw new ArgumentException("Choose a valid correction belonging to this orientation review.", nameof(proposal));
    }
    private static byte[] Hash(byte[] bytes) { using var hash = SHA256.Create(); return hash.ComputeHash(bytes); }
}
