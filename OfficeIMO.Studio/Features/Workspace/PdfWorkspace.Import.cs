using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal async Task<PdfImportPreparation?> PrepareImportAsync(IReadOnlyList<string> paths, CancellationToken token,
        Func<string, bool, CancellationToken, Task<string?>>? promptPassword = null) {
        ThrowIfDisposed();
        if (!CanImportPages) throw new InvalidOperationException("This document cannot safely import pages.");
        if (paths is null || paths.Count is < 1 or > 1000) throw new ArgumentException("Choose between 1 and 1000 PDFs to import.", nameof(paths));
        long revision = Revision;
        int targetCount = Pages.Count;
        string[] sources = paths.Select(ValidateSourcePdfPath).ToArray();
        var captured = new List<PdfImportSource>(sources.Length);
        long remainingBytes = Infrastructure.StudioStorageAccess.MaximumDocumentBytes;
        foreach (string source in sources) {
            token.ThrowIfCancellationRequested();
            if (remainingBytes == 0) throw new IOException("The selected PDFs exceed the 512 MiB import limit. Import fewer documents at a time.");
            var snapshot = await _storage.ReadSnapshotAsync(source, token, remainingBytes).ConfigureAwait(true);
            remainingBytes -= snapshot.Bytes.LongLength;
            string name = _storage.Describe(source).Name;
            string? password = null;
            while (true) {
                bool invalidPassword;
                try {
                    int count = await RunCancellableCpuWorkAsync(() => PdfDocument.Load(snapshot.Bytes,
                        new PdfLoadOptions { Password = password }).Inspect().PageCount, token).ConfigureAwait(true);
                    captured.Add(new(name, snapshot.Bytes, password, count));
                    break;
                } catch (PdfPasswordRequiredException) when (promptPassword is not null) { invalidPassword = false; }
                catch (PdfInvalidPasswordException) when (promptPassword is not null) { invalidPassword = true; }
                password = await promptPassword!(name, invalidPassword, token).ConfigureAwait(true);
                if (password is null) return null;
            }
        }
        return new(this, revision, targetCount, captured.ToArray());
    }

    internal async Task<int> ApplyImportAsync(PdfImportPreparation preparation, IReadOnlyList<PdfImportSelection> selections,
        int insertBefore, CancellationToken token, IProgress<PdfWorkspaceProgress>? progress = null) {
        ThrowIfDisposed();
        if (!ReferenceEquals(preparation.Owner, this) || preparation.Revision != Revision)
            throw new InvalidOperationException("The document changed after import preparation.");
        if (insertBefore < 1 || insertBefore > preparation.TargetPageCount + 1) throw new ArgumentOutOfRangeException(nameof(insertBefore));
        var selected = selections.Select(item => new PdfImportSelection(item.SourceIndex, (int[])item.Pages.Clone())).ToArray();
        if (selected.Length == 0 || selected.Select(item => item.SourceIndex).Distinct().Count() != selected.Length ||
            selected.Any(item => item.SourceIndex < 0 || item.SourceIndex >= preparation.Sources.Count || item.Pages.Length == 0))
            throw new ArgumentException("Choose distinct import sources and their pages.", nameof(selections));
        long imported = selected.Sum(item => (long)item.Pages.Length);
        if (imported > PdfImportPreparation.MaximumImportedPages) throw new ArgumentException("An import cannot exceed 100,000 selected pages.");
        string description = $"Imported {imported} pages from {selected.Length} PDF documents";
        await MutateAsync(PdfWorkspaceOperationKind.Import, description, [], document => {
            if (preparation.Revision != Revision) throw new InvalidOperationException("The document changed after import preparation.");
            int position = insertBefore;
            foreach (var selection in selected) {
                token.ThrowIfCancellationRequested();
                document = document.Pages.Insert(position, preparation.Sources[selection.SourceIndex].Select(selection.Pages));
                position += selection.Pages.Length;
            }
            return document;
        }, token, progress).ConfigureAwait(false);
        return (int)imported;
    }
}
