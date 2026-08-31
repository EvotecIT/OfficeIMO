using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;

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
    private byte[] _bytes;
    private string _baseFingerprint;
    private PdfDocumentInfo _documentInfo;
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
        PdfDocumentInfo documentInfo,
        PdfDocumentPreflight preflight,
        PdfWorkspaceRecoveryStore recoveryStore,
        string? recoveryPath) {
        Path = path;
        _bytes = bytes;
        _baseFingerprint = baseFingerprint;
        _documentInfo = documentInfo;
        _preflight = preflight;
        _recoveryStore = recoveryStore;
        RecoveryPath = recoveryPath;
    }

    internal event EventHandler? Changed;

    internal string Path { get; private set; }

    internal string FileName => System.IO.Path.GetFileName(Path);

    internal long FileSize => _bytes.LongLength;

    internal long Revision => _revision;

    internal PdfDocumentInfo DocumentInfo => _documentInfo;

    internal IReadOnlyList<PdfPageInfo> Pages => _documentInfo.Pages;

    internal IReadOnlyList<PdfWorkspaceOperation> Journal => _journal.AsReadOnly();

    internal bool IsDirty => _revision != _savedRevision;

    internal bool CanUndo => _undo.Count > 0;

    internal bool CanRedo => _redo.Count > 0;

    internal bool HasRecovery => !string.IsNullOrWhiteSpace(RecoveryPath) && File.Exists(RecoveryPath);

    internal string? RecoveryPath { get; private set; }

    internal bool CanMutatePages => _preflight.CanManipulatePages;

    internal bool CanExtractPages => CanPlan(PdfMutationOperation.ExtractPages);

    internal bool CanImportPages => CanPlan(PdfMutationOperation.MergeDocuments);

    internal bool CanEditAnnotations => CanPlan(PdfMutationOperation.ModifyAnnotations);

    internal bool CanEditPageContent => CanPlan(PdfMutationOperation.ModifyPageContent);

    internal bool CanRedact => CanPlan(PdfMutationOperation.Redact);

    internal bool CanFillForms => _preflight.CanFillSimpleFormFields || _preflight.CanAppendFormFieldRevision;

    internal bool CanFlattenForms => _preflight.CanFlattenSimpleFormFields;

    internal string? SecurityWarning {
        get {
            PdfDocumentSecurityInfo security = _documentInfo.Security;
            if (security.HasSignatures) {
                return "This PDF contains signatures. Page edits require a full rewrite and are disabled to avoid invalidating signed bytes.";
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
        PdfWorkspaceRecoveryStore? recoveryStore = null) {
        string fullPath = System.IO.Path.GetFullPath(path);
        if (!File.Exists(fullPath)) throw new FileNotFoundException("The selected PDF no longer exists.", fullPath);
        if (!string.Equals(System.IO.Path.GetExtension(fullPath), ".pdf", StringComparison.OrdinalIgnoreCase)) {
            throw new NotSupportedException("OfficeIMO Studio currently opens PDF documents.");
        }

        byte[] bytes = await File.ReadAllBytesAsync(fullPath, cancellationToken).ConfigureAwait(false);
        (PdfDocumentInfo Info, PdfDocumentPreflight Preflight) analysis = await Task.Run(
            () => Analyze(bytes),
            cancellationToken).ConfigureAwait(false);
        PdfWorkspaceRecoveryStore store = recoveryStore ?? new PdfWorkspaceRecoveryStore();
        string baseFingerprint = PdfWorkspaceRecoveryStore.Fingerprint(bytes);
        return new PdfWorkspace(
            fullPath,
            bytes,
            baseFingerprint,
            analysis.Info,
            analysis.Preflight,
            store,
            store.Find(fullPath, baseFingerprint));
    }

    internal PdfDocument CreateDocumentSnapshot() {
        ThrowIfDisposed();
        return PdfDocument.Open(_bytes);
    }

    internal byte[] CopyBytes() {
        ThrowIfDisposed();
        return (byte[])_bytes.Clone();
    }

    internal Task ReorderAsync(IReadOnlyList<int> pageNumbers, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAsync(PdfWorkspaceOperationKind.Reorder, "Reordered pages", pageNumbers, document => document.Pages.Reorder(pageNumbers.ToArray()), cancellationToken, progress);

    internal Task RotateAsync(IReadOnlyList<int> pageNumbers, int degrees, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAsync(PdfWorkspaceOperationKind.Rotate, $"Rotated {pageNumbers.Count} page(s) {degrees} degrees", pageNumbers, document => document.Pages.Rotate(degrees, pageNumbers.ToArray()), cancellationToken, progress);

    internal Task DeleteAsync(IReadOnlyList<int> pageNumbers, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAsync(PdfWorkspaceOperationKind.Delete, $"Deleted {pageNumbers.Count} page(s)", pageNumbers, document => document.Pages.Delete(pageNumbers.ToArray()), cancellationToken, progress);

    internal Task DuplicateAsync(IReadOnlyList<int> pageNumbers, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAsync(PdfWorkspaceOperationKind.Duplicate, $"Duplicated {pageNumbers.Count} page(s)", pageNumbers, document => document.Pages.Duplicate(pageNumbers.ToArray()), cancellationToken, progress);

    internal Task CropAsync(IReadOnlyList<int> pageNumbers, double left, double bottom, double right, double top, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAsync(PdfWorkspaceOperationKind.Crop, $"Cropped {pageNumbers.Count} page(s)", pageNumbers, document => document.Pages.CropAndTranslate(left, bottom, right, top, pageNumbers.ToArray()), cancellationToken, progress);

    internal Task CropByMarginAsync(IReadOnlyList<int> pageNumbers, double margin, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) {
        if (margin < 0D) throw new ArgumentOutOfRangeException(nameof(margin));
        PdfPageInfo[] pageInfo = pageNumbers.Select(pageNumber => Pages[pageNumber - 1]).ToArray();
        return MutateAsync(
            PdfWorkspaceOperationKind.Crop,
            $"Cropped {pageNumbers.Count} page(s) by {margin:0.#} pt",
            pageNumbers,
            document => {
                for (int index = 0; index < pageNumbers.Count; index++) {
                    PdfPageInfo info = pageInfo[index];
                    if ((margin * 2D) >= info.Width || (margin * 2D) >= info.Height) {
                        throw new InvalidOperationException($"The crop margin is too large for page {pageNumbers[index]}.");
                    }
                    document = document.Pages.CropAndTranslate(
                        margin,
                        margin,
                        info.Width - margin,
                        info.Height - margin,
                        pageNumbers[index]);
                }
                return document;
            },
            cancellationToken,
            progress);
    }

    internal Task InsertBlankAsync(int insertBeforePageNumber, double width, double height, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) {
        PdfDocument blank = PdfDocument.Create(compose => compose.Page(page => page.Size(width, height)));
        return MutateAsync(PdfWorkspaceOperationKind.InsertBlank, "Inserted a blank page", Array.Empty<int>(), document => document.Pages.Insert(insertBeforePageNumber, blank), cancellationToken, progress);
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
        return MutateBytesAsync(
            kind,
            description,
            new[] { gesture.PageNumber },
            bytes => PdfEditorCommandExecutor.Apply(bytes, PdfEditorCommandFactory.Create(bytes, tool, gesture, properties)),
            cancellationToken,
            progress);
    }

    internal async Task<PdfVerifiedRedactionResult> ApplyVerifiedRedactionAsync(
        PdfRedactionPlan reviewedPlan,
        long expectedRevision,
        string? removedTextMarker,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
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
                verified = PdfEditorCommandExecutor.ApplyVerifiedRedaction(bytes, reviewedPlan, removedTextMarker);
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
            PdfEditorCommand command = PdfEditorCommandFactory.Create(snapshot, PdfEditorTool.Redact, gesture, properties);
            return PdfEditorCommandExecutor.PlanRedaction(snapshot, command);
        }, cancellationToken);
    }

    internal Task FillFormFieldAsync(
        string fieldName,
        string value,
        bool flatten,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (string.IsNullOrWhiteSpace(fieldName)) throw new ArgumentException("Choose a form field.", nameof(fieldName));
        string normalizedName = fieldName.Trim();
        IReadOnlyDictionary<string, string> values = new Dictionary<string, string>(StringComparer.Ordinal) { [normalizedName] = value ?? string.Empty };
        return MutateBytesAsync(
            flatten ? PdfWorkspaceOperationKind.FormFlatten : PdfWorkspaceOperationKind.FormFill,
            flatten ? "Filled and flattened form field " + normalizedName : "Filled form field " + normalizedName,
            Array.Empty<int>(),
            bytes => {
                PdfDocument document = PdfDocument.Open(bytes);
                if (flatten) return document.Forms.FillAndFlatten(values).ToBytes();
                PdfMutationPlan plan = document.PlanMutation(PdfMutationOperation.FillFormFields, values.Keys);
                return (plan.ExecutionMode == PdfMutationExecutionMode.AppendOnly
                    ? document.Forms.AppendRevision(values)
                    : document.Forms.Fill(values)).ToBytes();
            },
            cancellationToken,
            progress);
    }

    internal Task FlattenFormFieldsAsync(
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(
            PdfWorkspaceOperationKind.FormFlatten,
            "Flattened form fields",
            Array.Empty<int>(),
            bytes => PdfDocument.Open(bytes).Forms.Flatten().ToBytes(),
            cancellationToken,
            progress);

    internal Task ApplyWatermarkAsync(
        string text,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (string.IsNullOrWhiteSpace(text)) throw new ArgumentException("Watermark text is required.", nameof(text));
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.Watermark,
            "Added text watermark",
            Enumerable.Range(1, Pages.Count).ToArray(),
            bytes => PdfDocument.Open(bytes).Stamp.TextWatermark(text.Trim(), new PdfTextStampOptions {
                FontSize = 42D,
                RotationDegrees = -35D,
                Color = PdfColor.FromRgb(148, 163, 184)
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
            bytes => PdfDocument.Open(bytes).Stamp.Content(
                (canvas, context) => canvas.Text(
                    context.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + " / " + context.PageCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    20D,
                    Math.Max(0D, context.Height - 28D),
                    Math.Max(1D, context.Width - 40D),
                    18D,
                    fontSize: 10D,
                    color: PdfColor.FromRgb(71, 84, 103),
                    align: PdfAlign.Center),
                new PdfCanvasStampOptions()).ToBytes(),
            cancellationToken,
            progress);

    internal Task UpdateAnnotationAsync(
        int objectNumber,
        string contents,
        string author,
        PdfColor color,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(
            PdfWorkspaceOperationKind.Annotation,
            "Updated annotation",
            Array.Empty<int>(),
            bytes => PdfDocument.Open(bytes).Annotations.Update(objectNumber, new PdfAnnotationUpdateOptions {
                Contents = contents ?? string.Empty,
                Title = author ?? string.Empty,
                Color = new[] { color.R, color.G, color.B },
                RegenerateAppearance = true
            }).Bytes,
            cancellationToken,
            progress);

    internal Task AddAnnotationReplyAsync(
        int objectNumber,
        string contents,
        string author,
        PdfColor color,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (string.IsNullOrWhiteSpace(contents)) throw new ArgumentException("Reply text is required.", nameof(contents));
        return MutateBytesAsync(
            PdfWorkspaceOperationKind.Annotation,
            "Added annotation reply",
            Array.Empty<int>(),
            bytes => PdfDocument.Open(bytes).Annotations.AddReply(objectNumber, contents.Trim(), new PdfAnnotationReplyOptions {
                Author = author,
                Color = new[] { color.R, color.G, color.B },
                CreatePopup = true
            }).Bytes,
            cancellationToken,
            progress);
    }

    internal Task FlattenAnnotationAsync(
        int objectNumber,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(
            PdfWorkspaceOperationKind.Annotation,
            "Flattened annotation",
            Array.Empty<int>(),
            bytes => PdfDocument.Open(bytes).Annotations.Flatten(new PdfAnnotationFlattenOptions { ObjectNumber = objectNumber }).Bytes,
            cancellationToken,
            progress);

    internal Task RemoveAnnotationAsync(
        int objectNumber,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(
            PdfWorkspaceOperationKind.Annotation,
            "Removed annotation",
            Array.Empty<int>(),
            bytes => PdfDocument.Open(bytes).Annotations.Remove(new PdfAnnotationRemovalOptions { ObjectNumber = objectNumber }).Bytes,
            cancellationToken,
            progress);

    internal async Task SaveAsync(string? path, CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) {
        ThrowIfDisposed();
        string destination = string.IsNullOrWhiteSpace(path) ? Path : System.IO.Path.GetFullPath(path);
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            string previousPath = Path;
            progress?.Report(new PdfWorkspaceProgress("Saving PDF", 0.2D));
            await PdfDocument.Open(_bytes).SaveAsync(destination, cancellationToken).ConfigureAwait(false);
            Path = destination;
            _baseFingerprint = PdfWorkspaceRecoveryStore.Fingerprint(_bytes);
            _savedRevision = _revision;
            _recoveryStore.Delete(previousPath);
            _recoveryStore.Delete(Path);
            RecoveryPath = null;
            progress?.Report(new PdfWorkspaceProgress("Saved", 1D));
            Changed?.Invoke(this, EventArgs.Empty);
        } finally {
            _operationGate.Release();
        }
    }

    internal Task UndoAsync(CancellationToken cancellationToken) => RestoreHistoryAsync(isUndo: true, cancellationToken);

    internal Task RedoAsync(CancellationToken cancellationToken) => RestoreHistoryAsync(isUndo: false, cancellationToken);

    internal async Task RestoreRecoveryAsync(CancellationToken cancellationToken) {
        ThrowIfDisposed();
        if (!HasRecovery || RecoveryPath is null) return;
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            byte[] recovered = await File.ReadAllBytesAsync(RecoveryPath, cancellationToken).ConfigureAwait(false);
            (PdfDocumentInfo Info, PdfDocumentPreflight Preflight) analysis = await Task.Run(
                () => Analyze(recovered),
                cancellationToken).ConfigureAwait(false);
            PushHistory(_undo, new Snapshot(_bytes, _revision));
            ClearHistory(_redo);
            _bytes = recovered;
            _documentInfo = analysis.Info;
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

    internal void DiscardRecovery() {
        ThrowIfDisposed();
        _recoveryStore.Delete(Path);
        RecoveryPath = null;
        Changed?.Invoke(this, EventArgs.Empty);
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
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
        if (!CanMutatePages) throw new InvalidOperationException(SecurityWarning ?? "This document cannot be safely rewritten.");
        await MutateBytesAsync(
            kind,
            description,
            pageNumbers,
            bytes => mutation(PdfDocument.Open(bytes)).ToBytes(),
            cancellationToken,
            progress).ConfigureAwait(false);
    }

    private async Task MutateBytesAsync(
        PdfWorkspaceOperationKind kind,
        string description,
        IReadOnlyList<int> pageNumbers,
        Func<byte[], byte[]> mutation,
        CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress) {
        ThrowIfDisposed();
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            progress?.Report(new PdfWorkspaceProgress(description, 0.1D));
            byte[] previousBytes = _bytes;
            byte[] candidateBytes = await RunCancellableCpuWorkAsync(
                () => mutation(previousBytes),
                cancellationToken).ConfigureAwait(false);
            progress?.Report(new PdfWorkspaceProgress("Validating changed document", 0.75D));
            (PdfDocumentInfo Info, PdfDocumentPreflight Preflight) candidateAnalysis = await RunCancellableCpuWorkAsync(
                () => Analyze(candidateBytes),
                cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();

            long nextRevision = ++_nextRevision;
            await _recoveryStore
                .WriteAsync(Path, _baseFingerprint, candidateBytes, nextRevision, cancellationToken)
                .ConfigureAwait(false);

            PushHistory(_undo, new Snapshot(previousBytes, _revision));
            ClearHistory(_redo);
            _bytes = candidateBytes;
            _documentInfo = candidateAnalysis.Info;
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

    private bool CanPlan(PdfMutationOperation operation) {
        try {
            return PdfDocument.Open(_bytes).PlanMutation(operation).CanExecute;
        } catch {
            return false;
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
            (PdfDocumentInfo Info, PdfDocumentPreflight Preflight) analysis = await Task.Run(
                () => Analyze(restore.Bytes),
                cancellationToken).ConfigureAwait(false);

            if (restore.Revision != _savedRevision) {
                await _recoveryStore
                    .WriteAsync(Path, _baseFingerprint, restore.Bytes, restore.Revision, cancellationToken)
                    .ConfigureAwait(false);
            }

            source.RemoveLast();
            _historyBytes -= restore.Bytes.LongLength;
            PushHistory(destination, new Snapshot(_bytes, _revision));
            _bytes = restore.Bytes;
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
                _recoveryStore.Delete(Path);
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

    private static (PdfDocumentInfo Info, PdfDocumentPreflight Preflight) Analyze(byte[] bytes) {
        PdfDocument document = PdfDocument.Open(bytes);
        return (document.Read.DocumentInfo(), document.Preflight());
    }

    private sealed record Snapshot(byte[] Bytes, long Revision);
}
