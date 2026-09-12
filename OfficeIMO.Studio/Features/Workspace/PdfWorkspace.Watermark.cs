using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal Task<IReadOnlyList<PdfWatermarkOptions>> ReadWatermarksAsync(CancellationToken cancellationToken) {
        ThrowIfDisposed();
        var snapshot = CreateDocumentSnapshot();
        return RunCancellableCpuWorkAsync(() => snapshot.Stamp.ReadWatermarks(), cancellationToken);
    }
    internal async Task<PdfWatermarkPreview> PrepareWatermarkAsync(PdfWatermarkOptions options, int previewPage,
        CancellationToken cancellationToken) {
        ThrowIfDisposed();
        PdfWatermarkOptions settings = options.Clone();
        PdfDocument source;
        long revision;
        await _operationGate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try {
            revision = _revision;
            source = CreateDocumentSnapshot();
        } finally { _operationGate.Release(); }
        return await RunCancellableCpuWorkAsync(() => {
            PdfDocument candidate = source.Stamp.Watermark(settings);
            cancellationToken.ThrowIfCancellationRequested();
            var pages = candidate.GetPageLayouts(options: null, cancellationToken);
            if (previewPage < 1 || previewPage > pages.Count)
                throw new ArgumentOutOfRangeException(nameof(previewPage));
            int[] selected = settings.TargetPages?.Resolve(pages.Count).ToArray() ?? Enumerable.Range(1, pages.Count).ToArray();
            if (!selected.Contains(previewPage))
                throw new ArgumentException("The preview page must receive the watermark.", nameof(previewPage));
            var page = pages[previewPage - 1];
            var rendered = candidate.Render.DisplayPage(previewPage, new PdfPageDisplayOptions {
                Scale = Math.Min(1.5D, 1000D / Math.Max(page.VisualWidth, page.VisualHeight)),
                MaximumOutputBytes = 8 * 1024 * 1024
            }, cancellationToken);
            if (!rendered.Succeeded || rendered.Bytes is null)
                throw new InvalidOperationException(string.Join(Environment.NewLine, rendered.Diagnostics));
            return new PdfWatermarkPreview(this, revision, candidate.ToBytes(), rendered.Bytes, selected, page.VisualWidth, page.VisualHeight);
        }, cancellationToken).ConfigureAwait(false);
    }

    internal Task ApplyWatermarkAsync(PdfWatermarkPreview preview, CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) {
        if (!ReferenceEquals(preview.Workspace, this))
            throw new InvalidOperationException("The watermark preview belongs to another document.");
        return MutateBytesAsync(PdfWorkspaceOperationKind.Watermark, "Added watermark", preview.Pages,
            _ => {
                if (_revision != preview.Revision)
                    throw new InvalidOperationException("The document changed. Preview the watermark again before applying it.");
                return preview.DocumentBytes;
            }, cancellationToken, progress);
    }
}

internal sealed record PdfWatermarkPreview(PdfWorkspace Workspace, long Revision, byte[] DocumentBytes,
    byte[] PageImage, int[] Pages, double PageWidth, double PageHeight);
