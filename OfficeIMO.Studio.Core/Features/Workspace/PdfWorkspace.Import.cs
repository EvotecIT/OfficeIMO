using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    /// <summary>
    /// Reads the PDFs to import. A protected source is opened with its entry in <paramref name="passwords"/> when given;
    /// otherwise <paramref name="promptPassword"/> asks for it, and without a prompt the preparation fails with a
    /// <see cref="PdfImportPasswordException"/> naming the source.
    /// </summary>
    internal async Task<PdfImportPreparation?> PrepareImportAsync(IReadOnlyList<string> paths, CancellationToken token,
        Func<string, bool, CancellationToken, Task<string?>>? promptPassword = null, IReadOnlyList<string?>? passwords = null) {
        ThrowIfDisposed();
        if (!CanImportPages) throw new InvalidOperationException("This document cannot safely import pages.");
        if (paths is null || paths.Count is < 1 or > 1000) throw new ArgumentException("Choose between 1 and 1000 PDFs to import.", nameof(paths));
        long revision = Revision;
        int targetCount = Pages.Count;
        string[] sources = paths.Select(ValidateSourcePdfPath).ToArray();
        var captured = new List<PdfImportSource>(sources.Length);
        long remainingBytes = Infrastructure.StudioDocumentStorage.MaximumDocumentBytes;
        for (int index = 0; index < sources.Length; index++) {
            string source = sources[index];
            token.ThrowIfCancellationRequested();
            if (remainingBytes == 0) throw new IOException("The selected PDFs exceed the 512 MiB import limit. Import fewer documents at a time.");
            var snapshot = await _storage.ReadSnapshotAsync(source, token,
                Math.Min(remainingBytes, StudioPdfSecurityPolicy.MaximumInputBytes)).ConfigureAwait(true);
            remainingBytes -= snapshot.Bytes.LongLength;
            string name = _storage.Describe(source).Name;
            string? password = passwords is not null && index < passwords.Count ? passwords[index] : null;
            while (true) {
                bool invalidPassword;
                try {
                    int count = await RunCancellableCpuWorkAsync(() => PdfDocument.Load(snapshot.Bytes,
                        StudioPdfSecurityPolicy.CreateLoadOptions(password)).Inspect().PageCount, token).ConfigureAwait(true);
                    captured.Add(new(name, snapshot.Bytes, password, count));
                    break;
                } catch (PdfPasswordRequiredException exception) {
                    if (promptPassword is null) throw new PdfImportPasswordException(index, name, invalid: false, exception);
                    invalidPassword = false;
                } catch (PdfInvalidPasswordException exception) {
                    if (promptPassword is null) throw new PdfImportPasswordException(index, name, invalid: true, exception);
                    invalidPassword = true;
                }
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
        if (imported < 1 || preparation.TargetPageCount + imported > PdfImportPreparation.MaximumImportedPages)
            throw new ArgumentException($"The resulting PDF cannot exceed {PdfImportPreparation.MaximumImportedPages:N0} pages.", nameof(selections));
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

/// <summary>A PDF chosen for import needs its password, or the one given was wrong.</summary>
internal sealed class PdfImportPasswordException(int sourceIndex, string name, bool invalid, Exception inner)
    : Exception(invalid ? $"The password for “{name}” is not correct." : $"“{name}” is protected by a password.", inner) {
    /// <summary>The source's position in the requested list, from zero.</summary>
    internal int SourceIndex { get; } = sourceIndex;
    internal string Name { get; } = name;
    internal bool Invalid { get; } = invalid;
}
