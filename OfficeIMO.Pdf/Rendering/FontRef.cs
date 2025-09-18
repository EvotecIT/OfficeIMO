namespace OfficeIMO.Pdf;

internal sealed class FontRef {
    public string Name { get; }
    public PdfStandardFont Font { get; }
    public int ObjectId { get; }
    public FontRef(string name, PdfStandardFont font, int objectId) { Name = name; Font = font; ObjectId = objectId; }
}

