using System.Security.Cryptography;
using System.Threading;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Ocr;

/// <summary>Immutable recognition proposals bound to exact source bytes, requiring explicitly reviewed values.</summary>
/// <remarks>Applying produces an in-memory document. Publication and host revision checks remain the caller's responsibility.</remarks>
public sealed class PdfFormOcrReview {
    private readonly byte[] _sourceBytes;
    private readonly PdfDocument _source;
    private readonly HashSet<PdfFormOcrProposal> _owned;
    private readonly PdfOcrMergeOptions _options;

    internal PdfFormOcrReview(PdfDocument source, PdfOcrMergeOptions options, PdfOcrMergeResult ocr,
        IReadOnlyList<PdfFormOcrProposal> proposals) {
        _sourceBytes = source.ToBytes(); _source = PdfDocument.Load(_sourceBytes, source.ReadOptions);
        _options = options.Clone(); Ocr = ocr; Proposals = proposals;
        _owned = new HashSet<PdfFormOcrProposal>(proposals);
        using var sha = SHA256.Create();
        SourceFingerprint = BitConverter.ToString(sha.ComputeHash(_sourceBytes)).Replace("-", "").ToLowerInvariant();
    }
    /// <summary>SHA-256 identity of the immutable source snapshot.</summary>
    public string SourceFingerprint { get; }
    /// <summary>Canonical OCR evidence, diagnostics and geometry.</summary>
    public PdfOcrMergeResult Ocr { get; }
    /// <summary>Proposals for existing fields, with no initial acceptance.</summary>
    public IReadOnlyList<PdfFormOcrProposal> Proposals { get; }

    /// <summary>Returns the captured visual page dimensions used by widget and word evidence.</summary>
    public (double Width, double Height) GetPageSize(int pageNumber) {
        var page = Ocr.NativeDocument.Pages.SingleOrDefault(page => page.PageNumber == pageNumber)
            ?? throw new ArgumentOutOfRangeException(nameof(pageNumber));
        return page.GetVisualPageSize();
    }

    /// <summary>Renders a captured source page for reviewing the value and its original geometry.</summary>
    public PdfPageRenderResult RenderPage(int pageNumber, CancellationToken cancellationToken = default) {
        if (!Ocr.Pages.Any(page => page.PageNumber == pageNumber)) throw new ArgumentOutOfRangeException(nameof(pageNumber));
        return PdfPageImageRenderer.RenderPages(_sourceBytes, PdfPageSelection.From(pageNumber), new() {
            Dpi = _options.Dpi, Format = PdfPageRenderFormat.Png, MaxPages = 1,
            MaxPixelsPerPage = _options.MaxPixelsPerPage, MaxOutputBytesPerPage = _options.MaxRenderedBytesPerPage,
            ImageCodec = _options.ImageCodec, ContinueOnError = false
        }, _source.ReadOptions, cancellationToken).Single();
    }

    /// <summary>Validates a reviewed correction using the canonical field metadata assessment.</summary>
    /// <remarks>Unknown script constraints, buttons and signatures cannot be overridden by acceptance.</remarks>
    public PdfFormFieldValueAssessment Assess(PdfFormOcrProposal proposal, PdfFormFieldValue value) {
        ValidateProposal(proposal);
        Guard.NotNull(value, nameof(value));
        return PdfFormFieldValueAssessment.Assess(proposal.Field, value);
    }

    /// <summary>Creates one atomic form-fill result from explicitly accepted and optionally corrected proposals.</summary>
    /// <remarks>Foreign proposals, stale bytes, unsupported constraints, invalid values and cancellation fail before output.
    /// Low-confidence or ambiguous evidence must be resolved by the human supplying the accepted value. An empty acceptance
    /// returns a source copy. Existing mutation planning enforces encryption, certification and field locks.
    /// Append-only filling retains the original byte prefix and generates updated normal appearance streams.</remarks>
    public PdfDocument Apply(PdfDocument currentDocument,
        IReadOnlyDictionary<PdfFormOcrProposal, PdfFormFieldValue> acceptedValues, CancellationToken cancellationToken = default) {
        Guard.NotNull(currentDocument, nameof(currentDocument));
        Guard.NotNull(acceptedValues, nameof(acceptedValues));
        cancellationToken.ThrowIfCancellationRequested();
        if (!_sourceBytes.SequenceEqual(currentDocument.ToBytes(cancellationToken)))
            throw new InvalidOperationException("The document changed after recognition. Recognize the current document again.");
        var values = new Dictionary<string, PdfFormFieldValue>(StringComparer.Ordinal);
        long characters = 0;
        foreach (var pair in acceptedValues) {
            cancellationToken.ThrowIfCancellationRequested();
            Guard.NotNull(pair.Value, nameof(acceptedValues));
            characters += pair.Value.Values.Sum(value => (long)value.Length);
            if (characters > _options.MaxOcrTextCharactersPerPage)
                throw PdfReadLimitException.Create(PdfReadLimitKind.OcrArtifacts, _options.MaxOcrTextCharactersPerPage, characters);
            var assessment = Assess(pair.Key, pair.Value);
            if (assessment.HasErrors) throw new ArgumentException(string.Join(" ", assessment.Issues.Where(issue => issue.IsError).Select(issue => issue.Message)), nameof(acceptedValues));
            values.Add(pair.Key.Field.Name!, PdfFormFieldValue.FromValues(pair.Value.Values));
        }
        if (values.Count == 0) return PdfDocument.Load(_sourceBytes, _source.ReadOptions);
        var plan = _source.PlanMutation(PdfMutationOperation.FillFormFields, values.Keys);
        cancellationToken.ThrowIfCancellationRequested();
        var result = plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly
            ? _source.Forms.AppendRevision(values, new PdfIncrementalFormFieldUpdateOptions {
                GenerateAppearanceStreams = true, KeepNeedAppearances = false
            }) : _source.Forms.Fill(values);
        cancellationToken.ThrowIfCancellationRequested();
        return result;
    }

    private void ValidateProposal(PdfFormOcrProposal proposal) {
        if (proposal is null || !_owned.Contains(proposal)) throw new ArgumentException("Choose a proposal from this recognition review.", nameof(proposal));
        if (!proposal.CanAccept) throw new InvalidOperationException(proposal.RejectionReason);
    }
}
