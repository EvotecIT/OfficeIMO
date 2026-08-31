namespace OfficeIMO.Studio.Features.Editor;

public sealed record PdfFormFieldViewModel(
    string Name,
    string Kind,
    string Value,
    bool IsReadOnly,
    IReadOnlyList<int> PageNumbers) {
    public string Label => Name + " · " + Kind + (IsReadOnly ? " · read only" : string.Empty);
}
