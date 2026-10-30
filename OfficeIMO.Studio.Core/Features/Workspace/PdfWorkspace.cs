using OfficeIMO.Pdf;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;
using System.Security.Cryptography;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Workspace;

/// <summary>
/// Reusable non-visual editing workspace over immutable OfficeIMO.Pdf snapshots. It owns dirty state,
/// bounded undo/redo, recovery, atomic saves, and operation progress; Avalonia only consumes this contract.
/// </summary>
internal sealed partial class PdfWorkspace : IDisposable {
    private const int MaximumHistoryEntries = 24;
    private const long MaximumHistoryBytes = 256L * 1024L * 1024L;
    private static readonly SemaphoreSlim ApplicationCpuWorkGate = new(1, 1);
    private readonly SemaphoreSlim _operationGate = new(1, 1);
    private readonly object _cpuWorkSync = new();
    private readonly LinkedList<Snapshot> _undo = new();
    private readonly LinkedList<Snapshot> _redo = new();
    private readonly List<PdfWorkspaceOperation> _journal = new();
    private readonly PdfWorkspaceRecoveryStore _recoveryStore;
    private readonly PdfLoadOptions _readOptions;
    private readonly StudioDocumentStorage _storage;
    private readonly Func<string, CancellationToken, ValueTask<bool>>? _canPublishOutput;
    private byte[] _bytes;
    private string _baseFingerprint;
    private string _sourceIdentityKey;
    private PdfDocumentViewInfo _documentInfo;
    private PdfDocumentPreflight _preflight;
    private long _historyBytes;
    private long _revision;
    private long _nextRevision;
    private long _savedRevision;
    private Task? _activeCpuWorker;
    private bool _disposed;

    private PdfWorkspace(
        string path,
        byte[] bytes,
        string baseFingerprint,
        string sourceIdentityKey,
        PdfDocumentViewInfo documentInfo,
        PdfDocumentPreflight preflight,
        PdfLoadOptions readOptions,
        PdfWorkspaceRecoveryStore recoveryStore,
        string? recoveryPath,
        Func<string, CancellationToken, ValueTask<bool>>? canPublishOutput,
        StudioDocumentStorage storage, string? recoveryIdentity) {
        Path = path;
        _bytes = bytes;
        _baseFingerprint = baseFingerprint;
        _sourceIdentityKey = sourceIdentityKey;
        _documentInfo = documentInfo;
        _preflight = preflight;
        _readOptions = readOptions;
        _annotationIdentities = BuildAnnotationIdentities(bytes, documentInfo);
        _recoveryStore = recoveryStore;
        RecoveryPath = recoveryPath;
        _canPublishOutput = canPublishOutput;
        _storage = storage;
        _recoveryIdentity = recoveryIdentity;
    }

    internal event EventHandler? Changed;

    internal string Path { get; private set; }

    internal string BaseFingerprint => _baseFingerprint;

    internal string FileName => _storage.Describe(Path).Name;
    internal bool UsesProviderPublication(string? destination = null) => _storage.UsesProviderPublication(destination ?? Path);

    internal long FileSize => _bytes.LongLength;

    internal long Revision => _revision;

    internal PdfDocumentInfo? DocumentInfo => _documentInfo.LogicalContent;

    internal PdfDocumentViewInfo ViewInfo => _documentInfo;

    internal IReadOnlyList<PdfPageInfo> Pages => _documentInfo.Pages;

    internal IReadOnlyList<PdfWorkspaceOperation> Journal => _journal.AsReadOnly();

    internal bool IsDirty => _revision != _savedRevision;

    internal bool CanUndo => _undo.Count > 0;

    internal bool CanRedo => _redo.Count > 0;

    internal bool HasRecovery => !string.IsNullOrWhiteSpace(RecoveryPath) && File.Exists(RecoveryPath);

    internal string? RecoveryPath { get; private set; }

    internal bool CanMutatePages => CanPlan(PdfMutationOperation.OrganizePages);

    internal bool CanRotatePages => CanPlan(PdfMutationOperation.RotatePages);

    internal bool CanCropPages => CanPlan(PdfMutationOperation.ModifyPageAttributes);

    internal bool CanResizePages => CanPlan(PdfMutationOperation.ModifyPageTree);

    internal bool CanInsertBlankPages => CanPlan(PdfMutationOperation.ModifyPageTree);

    internal bool CanExtractPages => CanPlan(PdfMutationOperation.ExtractPages);

    internal bool CanImportPages => CanPlan(PdfMutationOperation.ImportPages);

    internal bool CanEditAnnotations => CanPlan(PdfMutationOperation.ModifyAnnotations);

    internal bool CanEditPageContent => CanPlan(PdfMutationOperation.ModifyPageContent);

    internal bool CanRedact => CanPlan(PdfMutationOperation.Redact);

    internal bool HasForms => _preflight.Probe.HasForms;

    internal bool CanFillForms => CanPlan(PdfMutationOperation.FillFormFields);

    internal bool CanFlattenForms => _preflight.CanFlattenSimpleFormFields;

    internal bool CanAuthorForms => CanPlan(PdfMutationOperation.ModifyAcroForm);

    internal bool HasSignatures => _documentInfo.Security.HasSignatures;

    internal bool HasEncryption => _documentInfo.Security.HasEncryption;

    internal bool HasCertifiedRestrictions => _documentInfo.Security.HasDocMDPPermissions || _documentInfo.Security.HasUsageRights;

    internal string? SecurityWarning {
        get {
            PdfDocumentSecurityInfo security = _documentInfo.Security;
            if (security.HasSignatures) {
                return "This PDF contains signatures. Only actions permitted by its signature policy are available.";
            }
            if (security.HasEncryption) {
                return "This PDF is encrypted. Page edits are disabled unless OfficeIMO can safely preserve its security contract.";
            }
            if (security.HasDocMDPPermissions || security.HasUsageRights) {
                return "This PDF carries certification or usage-rights restrictions. Destructive page edits are disabled.";
            }
            if (!CanMutatePages) {
                return string.Join(" ", _preflight
                    .GetCapabilityDiagnostics(PdfPreflightCapability.ManipulatePages)
                    .Take(2));
            }
            return null;
        }
    }

    internal static async Task<PdfWorkspace> OpenAsync(
        string path,
        CancellationToken cancellationToken,
        PdfWorkspaceRecoveryStore? recoveryStore = null,
        string? password = null,
        Func<string, CancellationToken, ValueTask<bool>>? canPublishOutput = null,
        StudioDocumentStorage? storage = null, string? recoveryIdentity = null) {
        storage ??= new StudioDocumentStorage();
        string fullPath = OfficeStorageIdentity.Normalize(path);
        recoveryIdentity = recoveryIdentity is null ? null : OfficeStorageIdentity.Normalize(recoveryIdentity);
        if (recoveryIdentity?.Length > 4096) throw new ArgumentException("The recovery identity is too long.", nameof(recoveryIdentity));
        if (!string.Equals(System.IO.Path.GetExtension(storage.Describe(fullPath).Name), ".pdf", StringComparison.OrdinalIgnoreCase)) {
            throw new NotSupportedException("OfficeIMO Studio currently opens PDF documents.");
        }
        StudioStorageSnapshot source = await storage.ReadSnapshotAsync(fullPath, cancellationToken).ConfigureAwait(false);
        byte[] bytes = source.Bytes;
        string sourceIdentityKey = source.Identity;
        PdfLoadOptions readOptions = StudioPdfSecurityPolicy.CreateLoadOptions(password);
        (PdfDocumentViewInfo Info, PdfDocumentPreflight Preflight) analysis = await Task.Run(
            () => Analyze(bytes, readOptions),
            cancellationToken).ConfigureAwait(false);
        PdfWorkspaceRecoveryStore store = recoveryStore ?? new PdfWorkspaceRecoveryStore();
        string baseFingerprint = PdfWorkspaceRecoveryStore.Fingerprint(bytes);
        return new PdfWorkspace(
            fullPath,
            bytes,
            baseFingerprint,
            sourceIdentityKey,
            analysis.Info,
            analysis.Preflight,
            readOptions,
            store,
            store.Find(recoveryIdentity ?? fullPath, baseFingerprint), canPublishOutput, storage, recoveryIdentity);
    }

    internal PdfDocument CreateDocumentSnapshot() {
        ThrowIfDisposed();
        return LoadDocument(_bytes);
    }

    internal byte[] CopyBytes() {
        ThrowIfDisposed();
        return (byte[])_bytes.Clone();
    }

    /// <summary>The options the document was opened with, including its password, for reading copies of its bytes.</summary>
    internal PdfLoadOptions ReadOptions => _readOptions;

    internal Task ReorderAsync(IReadOnlyList<int> pageNumbers, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        ReorderWithIdentitiesAsync(pageNumbers, cancellationToken, progress);

    internal Task RotateAsync(IReadOnlyList<int> pageNumbers, int degrees, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAsync(PdfWorkspaceOperationKind.Rotate, $"Rotated {pageNumbers.Count} page(s) {degrees} degrees", pageNumbers, document => document.Pages.RotateBy(degrees, pageNumbers.ToArray()), cancellationToken, progress);

    internal Task DeleteAsync(IReadOnlyList<int> pageNumbers, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAsync(PdfWorkspaceOperationKind.Delete, $"Deleted {pageNumbers.Count} page(s)", pageNumbers, document => document.Pages.Delete(pageNumbers.ToArray()), cancellationToken, progress);

    internal Task DuplicateAsync(IReadOnlyList<int> pageNumbers, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAsync(PdfWorkspaceOperationKind.Duplicate, $"Duplicated {pageNumbers.Count} page(s)", pageNumbers, document => document.Pages.Duplicate(pageNumbers.ToArray()), cancellationToken, progress);

    internal Task CropAsync(IReadOnlyList<int> pageNumbers, double left, double bottom, double right, double top, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAsync(PdfWorkspaceOperationKind.Crop, $"Cropped {pageNumbers.Count} page(s)", pageNumbers, document => document.Pages.CropAndTranslate(left, bottom, right, top, pageNumbers.ToArray()), cancellationToken, progress);

    internal Task CropByMarginAsync(IReadOnlyList<int> pageNumbers, double margin, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) {
        if (!double.IsFinite(margin) || margin < 0D) throw new ArgumentOutOfRangeException(nameof(margin));
        return CropByMarginsAsync(pageNumbers, margin, margin, margin, margin, $"Cropped {pageNumbers.Count} page(s) by {margin:0.#} pt",
            cancellationToken, progress);
    }

    /// <summary>Crops inward from each displayed edge of the pages by its own margin in points, as one undoable change.</summary>
    internal Task CropByMarginsAsync(IReadOnlyList<int> pageNumbers, double left, double bottom, double right, double top,
        CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) {
        foreach (double margin in new[] { left, bottom, right, top })
            if (!double.IsFinite(margin) || margin < 0D) throw new ArgumentOutOfRangeException(nameof(left), "Crop margins are zero or more points.");
        return CropByMarginsAsync(pageNumbers, left, bottom, right, top, $"Cropped {pageNumbers.Count} page(s)", cancellationToken, progress);
    }

    private Task CropByMarginsAsync(IReadOnlyList<int> pageNumbers, double left, double bottom, double right, double top, string description,
        CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress) {
        PdfPageSelection selection = PdfPageSelection.From(pageNumbers.ToArray());
        return MutateAsync(
            PdfWorkspaceOperationKind.CropBox,
            description,
            pageNumbers,
            document => document.Pages.ApplyAttributes(PdfPageAttributeEdit.CropMargins(left, bottom, right, top, selection)),
            cancellationToken,
            progress);
    }

    internal Task ResetCropAsync(IReadOnlyList<int> pageNumbers, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAsync(PdfWorkspaceOperationKind.CropBox, $"Reset crop on {pageNumbers.Count} page(s)", pageNumbers,
            document => document.Pages.ApplyAttributes(PdfPageAttributeEdit.ResetCrop(PdfPageSelection.From(pageNumbers.ToArray()))),
            cancellationToken, progress);

    internal Task InsertBlankAsync(int insertBeforePageNumber, double width, double height, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) {
        // The blank has no external content to copy. Preserve the primary catalog at every insertion position,
        // including its boundaries, and demand the page-tree permission rather than external import permission.
        return MutateBytesAsync(PdfWorkspaceOperationKind.InsertBlank, "Inserted a blank page", Array.Empty<int>(),
            bytes => PdfTaggedPageOrganizer.InsertBlank(bytes, insertBeforePageNumber, width, height, _readOptions, cancellationToken), cancellationToken, progress);
    }

    internal Task ApplyEditorGestureAsync(
        PdfEditorTool tool,
        PdfEditorGesture gesture,
        PdfEditorProperties properties,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (tool is PdfEditorTool.Select or PdfEditorTool.Redact) throw new ArgumentException("This editor tool requires a different workflow.", nameof(tool));
        PdfWorkspaceOperationKind kind = tool is PdfEditorTool.AddText or PdfEditorTool.AddImage
            ? PdfWorkspaceOperationKind.AddedContent
            : PdfWorkspaceOperationKind.Annotation;
        string description = "Added " + GetToolDescription(tool) + " on page " + gesture.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture);
        if (kind == PdfWorkspaceOperationKind.Annotation) {
            return MutateAnnotationBytesAsync(kind, description, [gesture.PageNumber],
                bytes => PdfEditorCommandExecutor.ApplyAnnotation(LoadDocument(bytes), PdfEditorCommandFactory.Create(LoadDocument(bytes), tool, gesture, properties)),
                cancellationToken, progress);
        }
        return MutateBytesAsync(
            kind,
            description,
            new[] { gesture.PageNumber },
            bytes => PdfEditorCommandExecutor.Apply(LoadDocument(bytes), PdfEditorCommandFactory.Create(LoadDocument(bytes), tool, gesture, properties)),
            cancellationToken,
            progress);
    }

    internal async Task<PdfVerifiedRedactionResult> ApplyVerifiedRedactionAsync(
        PdfRedactionPlan reviewedPlan,
        long expectedRevision,
        string? removedTextMarker,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null,
        PdfSanitizationOptions? sanitizationOptions = null) {
        ArgumentNullException.ThrowIfNull(reviewedPlan);
        PdfVerifiedRedactionResult? verified = null;
        await MutateBytesAsync(
            PdfWorkspaceOperationKind.Redaction,
            "Applied and verified redaction on page " + reviewedPlan.Areas[0].PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture),
            reviewedPlan.Areas.Select(static area => area.PageNumber).Distinct().ToArray(),
            bytes => {
                if (_revision != expectedRevision) {
                    throw new InvalidOperationException("The document changed after this redaction was reviewed. Plan the redaction again before applying it.");
                }
                if (sanitizationOptions is not null) sanitizationOptions.CancellationToken = cancellationToken;
                verified = PdfEditorCommandExecutor.ApplyVerifiedRedaction(LoadDocument(bytes), reviewedPlan, removedTextMarker, sanitizationOptions, cancellationToken);
                return verified.Bytes;
            },
            cancellationToken,
            progress).ConfigureAwait(false);
        return verified ?? throw new InvalidOperationException("Redaction did not produce a verification result.");
    }

    internal Task<PdfRedactionPlan> PlanRedactionAsync(
        PdfEditorGesture gesture,
        PdfEditorProperties properties,
        CancellationToken cancellationToken) {
        ThrowIfDisposed();
        byte[] snapshot = CopyBytes();
        return Task.Run(() => {
            PdfEditorCommand command = PdfEditorCommandFactory.Create(LoadDocument(snapshot), PdfEditorTool.Redact, gesture, properties);
            return PdfEditorCommandExecutor.PlanRedaction(LoadDocument(snapshot), command);
        }, cancellationToken);
    }

    internal Task ApplyWatermarkAsync(
        string text,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (string.IsNullOrWhiteSpace(text)) throw new ArgumentException("Watermark text is required.", nameof(text));
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.Watermark,
            "Added text watermark",
            Enumerable.Range(1, Pages.Count).ToArray(),
            bytes => LoadDocument(bytes).Stamp.TextWatermark(text.Trim(), new PdfTextStampOptions {
                FontSize = 42D,
                RotationDegrees = -35D,
                Color = global::OfficeIMO.OfficeColor.FromRgb(148, 163, 184)
            }).ToBytes(),
            cancellationToken,
            progress);
    }

    internal Task ApplyPageNumbersAsync(
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(
            PdfWorkspaceOperationKind.PageNumbers,
            "Added page numbers",
            Enumerable.Range(1, Pages.Count).ToArray(),
            bytes => LoadDocument(bytes).Stamp.Content(
                (canvas, context) => canvas.Text(
                    context.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " / " + context.PageCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    global::OfficeIMO.OfficeLength.Points(20D),
                    global::OfficeIMO.OfficeLength.Points(Math.Max(0D, context.Height - 28D)),
                    global::OfficeIMO.OfficeLength.Points(Math.Max(1D, context.Width - 40D)),
                    global::OfficeIMO.OfficeLength.Points(18D),
                    fontSize: global::OfficeIMO.OfficeLength.Points(10D),
                    color: global::OfficeIMO.OfficeColor.FromRgb(71, 84, 103),
                    align: PdfAlign.Center),
                new PdfCanvasStampOptions()).ToBytes(),
            cancellationToken,
            progress);

    internal Task UpdateAnnotationAsync(
        int objectNumber,
        string? contents,
        string? author,
        global::OfficeIMO.OfficeColor? color,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAnnotationBytesAsync(
            PdfWorkspaceOperationKind.Annotation,
            "Updated annotation",
            Array.Empty<int>(),
            bytes => LoadDocument(bytes).Annotations.Update(objectNumber, new PdfAnnotationUpdateOptions {
                Contents = contents,
                Title = author,
                Color = color is { } value ? new[] { value.ToNormalizedRgb().R, value.ToNormalizedRgb().G, value.ToNormalizedRgb().B } : null,
                RegenerateAppearance = true
            }),
            cancellationToken,
            progress);

    internal Task AddAnnotationReplyAsync(
        int objectNumber,
        string contents,
        string author,
        global::OfficeIMO.OfficeColor color,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (string.IsNullOrWhiteSpace(contents)) throw new ArgumentException("Reply text is required.", nameof(contents));
        return MutateAnnotationBytesAsync(
            PdfWorkspaceOperationKind.Annotation,
            "Added annotation reply",
            Array.Empty<int>(),
            bytes => LoadDocument(bytes).Annotations.AddReply(objectNumber, contents.Trim(), new PdfAnnotationReplyOptions {
                Author = author,
                Color = new[] { color.ToNormalizedRgb().R, color.ToNormalizedRgb().G, color.ToNormalizedRgb().B },
                CreatePopup = true
            }),
            cancellationToken,
            progress);
    }

    internal Task FlattenAnnotationAsync(
        int objectNumber,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAnnotationBytesAsync(
            PdfWorkspaceOperationKind.Annotation,
            "Flattened annotation",
            Array.Empty<int>(),
            bytes => LoadDocument(bytes).Annotations.Flatten(new PdfAnnotationFlattenOptions { ObjectNumber = objectNumber }),
            cancellationToken,
            progress);

    internal Task FlattenAllAnnotationsAsync(
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAnnotationBytesAsync(
            PdfWorkspaceOperationKind.Annotation,
            "Flattened all supported annotations",
            Enumerable.Range(1, Pages.Count).ToArray(),
            bytes => LoadDocument(bytes).Annotations.Flatten(),
            cancellationToken,
            progress);

    internal Task RemoveAnnotationAsync(
        int objectNumber,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAnnotationBytesAsync(
            PdfWorkspaceOperationKind.Annotation,
            "Removed annotation",
            Array.Empty<int>(),
            bytes => LoadDocument(bytes).Annotations.Remove(new PdfAnnotationRemovalOptions { ObjectNumber = objectNumber }),
            cancellationToken,
            progress);

    internal Task UndoAsync(CancellationToken cancellationToken) => RestoreHistoryAsync(isUndo: true, cancellationToken);

    internal Task RedoAsync(CancellationToken cancellationToken) => RestoreHistoryAsync(isUndo: false, cancellationToken);

    internal async Task RestoreRecoveryAsync(CancellationToken cancellationToken) {
        ThrowIfDisposed();
        if (!HasRecovery || RecoveryPath is null) return;
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            byte[] recovered = await Task.Run(
                () => _recoveryStore.ReadVerifiedSnapshot(RecoveryIdentity, RecoveryBaseFingerprint)
                    ?? throw new InvalidDataException("The recovery snapshot is no longer valid. The open document has not been changed."),
                cancellationToken).ConfigureAwait(false);
            (PdfDocumentViewInfo Info, PdfDocumentPreflight Preflight) analysis = await Task.Run(
                () => Analyze(recovered, _readOptions),
                cancellationToken).ConfigureAwait(false);
            if (_documentInfo.Security.HasEncryption && !analysis.Info.Security.HasEncryption)
                throw new InvalidDataException("The recovery snapshot removes the open document's encryption.");
            PushHistory(_undo, new Snapshot(_bytes, _revision, _annotationIdentities));
            ClearHistory(_redo);
            _bytes = recovered;
            _documentInfo = analysis.Info;
            _annotationIdentities = BuildAnnotationIdentities(recovered, analysis.Info);
            _preflight = analysis.Preflight;
            _revision = ++_nextRevision;
            TrimHistory();
            RecoveryPath = null;
            _journal.Add(new PdfWorkspaceOperation(
                _revision,
                PdfWorkspaceOperationKind.RecoveryRestore,
                "Restored recovered edits",
                Array.Empty<int>(),
                DateTimeOffset.UtcNow));
            Changed?.Invoke(this, EventArgs.Empty);
        } finally {
            _operationGate.Release();
        }
    }

    internal async Task DiscardRecoveryAsync(CancellationToken cancellationToken = default) {
        ThrowIfDisposed();
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            await _recoveryStore.DeleteAsync(RecoveryIdentity, cancellationToken).ConfigureAwait(false);
            RecoveryPath = null;
            Changed?.Invoke(this, EventArgs.Empty);
        } finally { _operationGate.Release(); }
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        ClearCapabilityCache();
        _operationGate.Dispose();
        _undo.Clear();
        _redo.Clear();
        _journal.Clear();
    }

    private async Task MutateAsync(
        PdfWorkspaceOperationKind kind,
        string description,
        IReadOnlyList<int> pageNumbers,
        Func<PdfDocument, PdfDocument> mutation,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress) {
        ThrowIfDisposed();
        // Every mutation delegates authorization and source preservation to its owning PDF operation.
        await MutateBytesAsync(
            kind,
            description,
            pageNumbers,
            bytes => mutation(LoadDocument(bytes)).ToBytes(),
            cancellationToken,
            progress).ConfigureAwait(false);
    }

    private async Task MutateBytesAsync(
        PdfWorkspaceOperationKind kind,
        string description,
        IReadOnlyList<int> pageNumbers,
        Func<byte[], byte[]> mutation,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress,
        bool detachCpuWorkOnCancellation = true,
        Func<AnnotationMapping?>? getAnnotationMapping = null) {
        ThrowIfDisposed();
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            progress?.Report(new PdfWorkspaceProgress(description, 0.1D));
            byte[] previousBytes = _bytes;
            byte[] candidateBytes = detachCpuWorkOnCancellation
                ? await RunCancellableCpuWorkAsync(
                    () => mutation(previousBytes),
                    cancellationToken).ConfigureAwait(false)
                : await RunNonDetachableCpuWorkAsync(
                    () => mutation(previousBytes),
                    cancellationToken).ConfigureAwait(false);
            progress?.Report(new PdfWorkspaceProgress("Validating changed document", 0.75D));
            (PdfDocumentViewInfo Info, PdfDocumentPreflight Preflight) candidateAnalysis = await RunCancellableCpuWorkAsync(
                () => Analyze(candidateBytes, _readOptions),
                cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (_documentInfo.Security.HasEncryption && !candidateAnalysis.Info.Security.HasEncryption)
                throw new InvalidOperationException("This edit would remove PDF encryption. Use the owner-authorized decrypted-copy workflow instead.");

            var nextIdentities = BuildAnnotationIdentities(candidateBytes, candidateAnalysis.Info, _annotationIdentities, getAnnotationMapping?.Invoke());
            long nextRevision = ++_nextRevision;
            await WriteRecoveryAsync(candidateBytes, nextRevision, cancellationToken)
                .ConfigureAwait(false);

            PushHistory(_undo, new Snapshot(previousBytes, _revision, _annotationIdentities));
            ClearHistory(_redo);
            _bytes = candidateBytes;
            _documentInfo = candidateAnalysis.Info;
            _annotationIdentities = nextIdentities;
            _preflight = candidateAnalysis.Preflight;
            _revision = nextRevision;
            _journal.Add(new PdfWorkspaceOperation(_revision, kind, description, pageNumbers.ToArray(), DateTimeOffset.UtcNow));
            TrimHistory();
            progress?.Report(new PdfWorkspaceProgress("Edit complete", 1D));
            Changed?.Invoke(this, EventArgs.Empty);
        } finally {
            _operationGate.Release();
        }
    }

    private static string GetToolDescription(PdfEditorTool tool) => tool switch {
        PdfEditorTool.Note => "note",
        PdfEditorTool.FreeText => "free text",
        PdfEditorTool.Highlight => "highlight",
        PdfEditorTool.Underline => "underline",
        PdfEditorTool.StrikeOut => "strikeout",
        PdfEditorTool.Rectangle => "rectangle",
        PdfEditorTool.Ellipse => "ellipse",
        PdfEditorTool.Line => "line",
        PdfEditorTool.Ink => "ink",
        PdfEditorTool.Squiggly => "squiggly underline",
        PdfEditorTool.Polygon => "polygon",
        PdfEditorTool.Stamp => "stamp",
        PdfEditorTool.AddText => "text",
        PdfEditorTool.AddImage => "image",
        PdfEditorTool.Link => "link",
        PdfEditorTool.SignatureAppearance => "visual signature appearance",
        _ => throw new ArgumentOutOfRangeException(nameof(tool), tool, "Unsupported editor tool.")
    };

    private async Task RestoreHistoryAsync(bool isUndo, CancellationToken cancellationToken) {
        ThrowIfDisposed();
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            LinkedList<Snapshot> source = isUndo ? _undo : _redo;
            LinkedList<Snapshot> destination = isUndo ? _redo : _undo;
            if (source.Last is null) return;

            Snapshot restore = source.Last.Value;
            (PdfDocumentViewInfo Info, PdfDocumentPreflight Preflight) analysis = await Task.Run(
                () => Analyze(restore.Bytes, _readOptions),
                cancellationToken).ConfigureAwait(false);

            if (restore.Revision != _savedRevision || _preparedExternalSaves.Count > 0) {
                await WriteRecoveryAsync(restore.Bytes, restore.Revision, cancellationToken)
                    .ConfigureAwait(false);
            } else {
                // Keep history and in-memory edits intact if cleanup fails or is canceled.
                await _recoveryStore.DeleteAsync(RecoveryIdentity, cancellationToken).ConfigureAwait(false);
            }

            source.RemoveLast();
            _historyBytes -= restore.Bytes.LongLength;
            PushHistory(destination, new Snapshot(_bytes, _revision, _annotationIdentities));
            _bytes = restore.Bytes;
            _annotationIdentities = restore.AnnotationIdentities;
            _revision = restore.Revision;
            _documentInfo = analysis.Info;
            _preflight = analysis.Preflight;
            _journal.Add(new PdfWorkspaceOperation(
                _revision,
                isUndo ? PdfWorkspaceOperationKind.Undo : PdfWorkspaceOperationKind.Redo,
                isUndo ? "Undo" : "Redo",
                Array.Empty<int>(),
                DateTimeOffset.UtcNow));
            if (!IsDirty) {
                RecoveryPath = null;
            }
            TrimHistory();
            Changed?.Invoke(this, EventArgs.Empty);
        } finally {
            _operationGate.Release();
        }
    }

    private void PushHistory(LinkedList<Snapshot> history, Snapshot snapshot) {
        history.AddLast(snapshot);
        _historyBytes += snapshot.Bytes.LongLength;
    }

    private void ClearHistory(LinkedList<Snapshot> history) {
        foreach (Snapshot snapshot in history) _historyBytes -= snapshot.Bytes.LongLength;
        history.Clear();
    }

    private void TrimHistory() {
        while ((_undo.Count + _redo.Count) > MaximumHistoryEntries || _historyBytes > MaximumHistoryBytes) {
            LinkedList<Snapshot> target = _undo.First is not null ? _undo : _redo;
            if (target.First is null) break;
            _historyBytes -= target.First.Value.Bytes.LongLength;
            target.RemoveFirst();
        }
    }

    private void ThrowIfDisposed() => ObjectDisposedException.ThrowIf(_disposed, this);

    private PdfDocument LoadDocument(byte[] bytes) => PdfDocument.Load(bytes, _readOptions);

    private static (PdfDocumentViewInfo Info, PdfDocumentPreflight Preflight) Analyze(byte[] bytes, PdfLoadOptions readOptions) {
        PdfDocument document = PdfDocument.Load(bytes, readOptions);
        return (document.InspectForViewing(), document.Preflight());
    }

    private sealed record Snapshot(byte[] Bytes, long Revision, IReadOnlyDictionary<int, Guid> AnnotationIdentities);
}
