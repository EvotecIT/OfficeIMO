using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

/// <summary>Captured import inputs; review and application consume the same immutable PDF bytes.</summary>
internal sealed class PdfImportPreparation(PdfWorkspace owner, long revision, int targetPageCount, PdfImportSource[] sources) {
    internal const int MaximumImportedPages = 100_000;
    internal PdfWorkspace Owner { get; } = owner;
    internal long Revision { get; } = revision;
    internal int TargetPageCount { get; } = targetPageCount;
    internal IReadOnlyList<PdfImportSource> Sources { get; } = Array.AsReadOnly(sources);
}

internal sealed class PdfImportSource(string name, byte[] bytes, string? password, int pageCount) {
    internal string Name { get; } = name;
    internal int PageCount { get; } = pageCount;
    internal PdfDocument Select(int[] pages) {
        if (pages.Length == 0 || pages.Any(page => page < 1 || page > PageCount))
            throw new ArgumentException("Select valid source pages to import.");
        PdfDocument document = PdfDocument.Load(bytes, new PdfLoadOptions { Password = password });
        return pages.SequenceEqual(Enumerable.Range(1, PageCount)) ? document : document.Pages.Extract(pages);
    }
}

internal sealed record PdfImportSelection(int SourceIndex, int[] Pages);
