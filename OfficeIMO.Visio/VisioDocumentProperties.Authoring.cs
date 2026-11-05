namespace OfficeIMO.Visio;
public partial class VisioDocumentProperties {
    /// <summary>Sets the document title.</summary>
    public VisioDocumentProperties SetTitle(string title) { Title = title; return this; }
    /// <summary>Sets the document author.</summary>
    public VisioDocumentProperties SetAuthor(string author) { Author = author; return this; }
}
