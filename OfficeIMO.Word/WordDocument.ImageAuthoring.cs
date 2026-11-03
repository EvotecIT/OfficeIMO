namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Adds an image in a new paragraph and returns its editable occurrence.</summary>
    public WordImage AddImage(string path, OfficeLength? width = null, OfficeLength? height = null, WordImageOptions? options = null) {
        ArgumentNullException.ThrowIfNull(path);
        var resolved = options ?? new WordImageOptions();
        return AddParagraph().AddImage(path, width, height, resolved);
    }

    /// <summary>Adds an image stream in a new paragraph without disposing the caller's stream.</summary>
    public WordImage AddImage(Stream stream, OfficeLength? width = null, OfficeLength? height = null, WordImageOptions? options = null) {
        ArgumentNullException.ThrowIfNull(stream);
        var resolved = options ?? new WordImageOptions();
        return AddParagraph().AddImage(stream, resolved.FileName, width, height, resolved);
    }
}
