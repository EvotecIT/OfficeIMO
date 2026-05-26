namespace OfficeIMO.Pdf;

internal sealed class BookmarkBlock : IPdfBlock {
    public string Name { get; }

    public BookmarkBlock(string name) {
        Guard.NotNullOrWhiteSpace(name, nameof(name));
        Name = name;
    }
}
