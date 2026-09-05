namespace OfficeIMO.Studio.Features.Reader;

public sealed record PdfBookmarkViewModel(string Title, int Level, int? PageNumber) {
    public string IndentedTitle => new string(' ', Math.Max(0, Level - 1) * 2) + Title;
}
