using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Visual evidence that supported a proposed static-form field.</summary>
public enum PdfStaticFormEvidenceKind {
    /// <summary>A stroked, empty rectangular field outline.</summary>
    OutlinedField,
    /// <summary>A horizontal writing line.</summary>
    Underline,
    /// <summary>A small, square checkbox outline.</summary>
    CheckBox
}

/// <summary>Optional, caller-supplied OCR text in top-left visual page coordinates.</summary>
public sealed class PdfStaticFormTextEvidence {
    /// <summary>Creates positioned text evidence without taking ownership of an OCR provider.</summary>
    public PdfStaticFormTextEvidence(int pageNumber, string text, double left, double top, double right, double bottom, double confidence) {
        Guard.PositiveInteger(pageNumber, nameof(pageNumber));
        Guard.NotNullOrWhiteSpace(text, nameof(text));
        if (text.Length > 256) throw new ArgumentOutOfRangeException(nameof(text), "OCR evidence text must not exceed 256 characters.");
        if (!IsFinite(left) || !IsFinite(top) || !IsFinite(right) || !IsFinite(bottom) ||
            left < 0D || top < 0D || right <= left || bottom <= top) {
            throw new ArgumentOutOfRangeException(nameof(left), "OCR evidence requires a finite positive visual rectangle.");
        }
        if (!IsFinite(confidence) || confidence < 0D || confidence > 1D) throw new ArgumentOutOfRangeException(nameof(confidence));
        PageNumber = pageNumber;
        Text = text;
        Left = left;
        Top = top;
        Right = right;
        Bottom = bottom;
        Confidence = confidence;
    }

    /// <summary>One-based source page number.</summary>
    public int PageNumber { get; }
    /// <summary>Recognized text.</summary>
    public string Text { get; }
    /// <summary>Left visual coordinate.</summary>
    public double Left { get; }
    /// <summary>Top visual coordinate.</summary>
    public double Top { get; }
    /// <summary>Right visual coordinate.</summary>
    public double Right { get; }
    /// <summary>Bottom visual coordinate.</summary>
    public double Bottom { get; }
    /// <summary>Provider confidence from zero through one.</summary>
    public double Confidence { get; }

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}

/// <summary>Limits and page selection for reviewable static-form recognition.</summary>
public sealed class PdfStaticFormRecognitionOptions {
    /// <summary>Pages to inspect; null inspects every page within the page limit.</summary>
    public PdfPageSelection? PageSelection { get; set; }
    /// <summary>Maximum selected pages.</summary>
    public int MaxPages { get; set; } = 50;
    /// <summary>Maximum proposals retained across the document.</summary>
    public int MaxProposals { get; set; } = 250;
    /// <summary>Maximum primitive comparisons spent evaluating visual field candidates.</summary>
    public int MaxCandidateScanWork { get; set; } = 1_000_000;
    /// <summary>Maximum diagnostics retained, including a truncation marker when further diagnostics are omitted.</summary>
    public int MaxDiagnostics { get; set; } = 1_000;
    /// <summary>Maximum supplied OCR text items.</summary>
    public int MaxOcrTextItems { get; set; } = 5_000;
    /// <summary>Minimum confidence retained as a proposal.</summary>
    public double MinimumConfidence { get; set; } = 0.6D;

    internal void Validate() {
        if (MaxPages <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPages));
        if (MaxProposals <= 0) throw new ArgumentOutOfRangeException(nameof(MaxProposals));
        if (MaxCandidateScanWork <= 0) throw new ArgumentOutOfRangeException(nameof(MaxCandidateScanWork));
        if (MaxDiagnostics <= 0) throw new ArgumentOutOfRangeException(nameof(MaxDiagnostics));
        if (MaxOcrTextItems < 0) throw new ArgumentOutOfRangeException(nameof(MaxOcrTextItems));
        if (double.IsNaN(MinimumConfidence) || MinimumConfidence < 0D || MinimumConfidence > 1D) {
            throw new ArgumentOutOfRangeException(nameof(MinimumConfidence));
        }
    }
}

/// <summary>A field candidate that must be reviewed before authoring.</summary>
public sealed class PdfStaticFormFieldProposal {
    internal PdfStaticFormFieldProposal(int index, int pageNumber, int suggestedTabIndex, string label, string suggestedName, PdfFormFieldCreationKind kind, PdfPageRectangle rectangle, PdfLogicalVisualBounds visualBounds, double confidence, PdfStaticFormEvidenceKind evidenceKind, bool usedOcrLabel) {
        Index = index;
        PageNumber = pageNumber;
        SuggestedTabIndex = suggestedTabIndex;
        Label = label;
        SuggestedName = suggestedName;
        Kind = kind;
        Rectangle = rectangle;
        VisualBounds = visualBounds;
        Confidence = confidence;
        EvidenceKind = evidenceKind;
        UsedOcrLabel = usedOcrLabel;
    }

    /// <summary>Stable index within this report, used for explicit approval.</summary>
    public int Index { get; }
    /// <summary>One-based source page.</summary>
    public int PageNumber { get; }
    /// <summary>One-based visual reading-order suggestion within the page; callers must review existing widgets before applying a page tab order.</summary>
    public int SuggestedTabIndex { get; }
    /// <summary>Nearby native or caller-supplied OCR label.</summary>
    public string Label { get; }
    /// <summary>Unique suggested AcroForm field name.</summary>
    public string SuggestedName { get; }
    /// <summary>Conservatively inferred field kind: text or checkbox.</summary>
    public PdfFormFieldCreationKind Kind { get; }
    /// <summary>Proposed widget rectangle in PDF default user space.</summary>
    public PdfPageRectangle Rectangle { get; }
    /// <summary>Proposed widget rectangle in top-left visual page coordinates.</summary>
    public PdfLogicalVisualBounds VisualBounds { get; }
    /// <summary>Review confidence from zero through one, not a correctness guarantee.</summary>
    public double Confidence { get; }
    /// <summary>Visual feature that supplied the field geometry.</summary>
    public PdfStaticFormEvidenceKind EvidenceKind { get; }
    /// <summary>Whether the matched label came from caller-supplied OCR evidence.</summary>
    public bool UsedOcrLabel { get; }

    /// <summary>Creates editable field options from this proposal for an explicit caller transaction.</summary>
    public PdfFormFieldCreateOptions ToCreateOptions() => new PdfFormFieldCreateOptions {
        Name = SuggestedName,
        Kind = Kind,
        PageNumber = PageNumber,
        X = Rectangle.Left,
        Y = Rectangle.Bottom,
        Width = Rectangle.Width,
        Height = Rectangle.Height,
        Style = new PdfFormFieldStyle {
            AlternateName = Label,
            BackgroundColor = null,
            BorderColor = null,
            BorderWidth = 0D
        }
    };
}

/// <summary>Reason a plausible visual candidate was not proposed.</summary>
public sealed class PdfStaticFormRecognitionDiagnostic {
    internal PdfStaticFormRecognitionDiagnostic(string code, int pageNumber, string message) {
        Code = code;
        PageNumber = pageNumber;
        Message = message;
    }
    /// <summary>Stable machine-readable reason.</summary>
    public string Code { get; }
    /// <summary>One-based source page.</summary>
    public int PageNumber { get; }
    /// <summary>Human-readable review detail.</summary>
    public string Message { get; }
}

/// <summary>Source-bound proposals and diagnostics; analysis never modifies the PDF.</summary>
public sealed class PdfStaticFormRecognitionReport {
    private readonly byte[] _analyzedPdf;
    private readonly PdfLoadOptions _readOptions;
    private readonly string _sourceSha256;
    internal PdfStaticFormRecognitionReport(byte[] analyzedPdf, PdfLoadOptions readOptions, IReadOnlyList<PdfStaticFormFieldProposal> proposals, IReadOnlyList<PdfStaticFormRecognitionDiagnostic> diagnostics) {
        _analyzedPdf = (byte[])analyzedPdf.Clone();
        _readOptions = readOptions;
        _sourceSha256 = PdfArtifactFingerprint.ComputeSha256(_analyzedPdf);
        Proposals = Array.AsReadOnly(proposals.ToArray());
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
    }
    /// <summary>Proposed fields in page and reading order.</summary>
    public IReadOnlyList<PdfStaticFormFieldProposal> Proposals { get; }
    /// <summary>Collision, unsupported, and low-confidence evidence.</summary>
    public IReadOnlyList<PdfStaticFormRecognitionDiagnostic> Diagnostics { get; }
    /// <summary>SHA-256 of the exact analyzed source artifact.</summary>
    public string SourceSha256 => _sourceSha256;

    /// <summary>Creates only explicitly selected proposals against the immutable PDF snapshot analyzed by this report.</summary>
    public PdfAcroFormEditResult ApplySelected(IReadOnlyCollection<int> proposalIndices, CancellationToken cancellationToken = default) {
        Guard.NotNull(proposalIndices, nameof(proposalIndices));
        cancellationToken.ThrowIfCancellationRequested();
        if (proposalIndices.Count == 0) throw new ArgumentException("Select at least one proposal to create.", nameof(proposalIndices));
        int[] selected = proposalIndices.Distinct().OrderBy(static index => index).ToArray();
        if (selected.Length != proposalIndices.Count || selected.Any(index => index < 0 || index >= Proposals.Count)) {
            throw new ArgumentOutOfRangeException(nameof(proposalIndices), "Proposal indices must be distinct and present in this report.");
        }
        return PdfAcroFormEditor.Edit(_analyzedPdf, edit => {
            foreach (int index in selected) {
                cancellationToken.ThrowIfCancellationRequested();
                edit.Create(Proposals[index].ToCreateOptions());
            }
        }, _readOptions, cancellationToken: cancellationToken);
    }
}
