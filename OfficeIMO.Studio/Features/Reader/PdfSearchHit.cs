namespace OfficeIMO.Studio.Features.Reader;

public sealed record PdfSearchHit(int PageNumber, string Snippet) {
    public string Label => $"Page {PageNumber}: {Snippet}";
}
