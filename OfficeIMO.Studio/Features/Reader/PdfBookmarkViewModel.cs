namespace OfficeIMO.Studio.Features.Reader;

/// <summary>
/// One outline entry. <paramref name="Id"/> matches the engine's bookmark edit session id for entries that
/// resolve to a page; unresolved entries are shown but cannot be edited.
/// </summary>
public sealed record PdfBookmarkViewModel(string Title, int Level, int? PageNumber, string? Id = null, string? ParentId = null, int Index = 0) {
    public string IndentedTitle => new string(' ', Math.Max(0, Level - 1) * 2) + Title;
    public double Indent => Math.Max(0, Level - 1) * 14D;
    public string PageLabel => PageNumber is int page ? page.ToString(System.Globalization.CultureInfo.CurrentCulture) : string.Empty;
    public bool IsEditable => Id is not null;
}
