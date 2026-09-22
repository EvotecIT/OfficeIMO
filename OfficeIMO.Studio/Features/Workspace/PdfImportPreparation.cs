using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Workspace;

/// <summary>Captured import inputs; review and application consume the same immutable PDF bytes.</summary>
internal sealed class PdfImportPreparation(PdfWorkspace owner, long revision, int targetPageCount, PdfImportSource[] sources) {
    internal const int MaximumImportedPages = Reader.StudioPdfSecurityPolicy.MaximumPages;
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
        PdfDocument document = PdfDocument.Load(bytes, StudioPdfSecurityPolicy.CreateLoadOptions(password));
        if (document.Inspect().PageCount != PageCount)
            throw new InvalidDataException("The captured import source page count changed after preparation.");
        return pages.SequenceEqual(Enumerable.Range(1, PageCount)) ? document : document.Pages.Extract(pages);
    }
}

internal sealed record PdfImportSelection(int SourceIndex, int[] Pages);
