namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Adds an image in a new paragraph and returns its editable occurrence.</summary>
    public WordImage AddImage(string path, OfficeLength? width = null, OfficeLength? height = null, WordImageOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        var resolved = options ?? new WordImageOptions();
        return resolved.Apply(AddParagraph().AddImage(path, width?.ToPixels(), height?.ToPixels(), resolved.Wrapping, resolved.Description));
    }

    /// <summary>Adds an image stream in a new paragraph without disposing the caller's stream.</summary>
    public WordImage AddImage(Stream stream, OfficeLength? width = null, OfficeLength? height = null, WordImageOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        var resolved = options ?? new WordImageOptions();
        return resolved.Apply(AddParagraph().AddImage(stream, resolved.FileName, width?.ToPixels(), height?.ToPixels(), resolved.Wrapping, resolved.Description));
    }
}
