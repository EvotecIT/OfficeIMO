namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Embeds an image file and returns its editable occurrence.</summary>
    public WordImage AddImage(string path, OfficeLength? width = null, OfficeLength? height = null,
        WordImageOptions? options = null) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        var resolved = options ?? new WordImageOptions();
        if (resolved.Vml) {
            using var stream = File.OpenRead(path);
            return resolved.Apply(InsertImageVml(stream, Path.GetFileName(path), width?.ToPoints(), height?.ToPoints(), resolved.Description));
        }
        return AttachImage(new WordImage(_document, this, path, width?.ToPixels(), height?.ToPixels(), resolved.Wrapping, resolved.Description), resolved);
    }

    /// <summary>Embeds an image stream without disposing the caller's stream.</summary>
    public WordImage AddImage(Stream stream, string fileName, OfficeLength? width = null, OfficeLength? height = null,
        WordImageOptions? options = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        var resolved = options ?? new WordImageOptions();
        if (resolved.Vml) return resolved.Apply(InsertImageVml(stream, fileName, width?.ToPoints(), height?.ToPoints(), resolved.Description));
        return AttachImage(new WordImage(_document, this, stream, fileName, width?.ToPixels(), height?.ToPixels(), resolved.Wrapping, resolved.Description), resolved);
    }

    /// <summary>Adds a linked external image. The image remains an external relationship.</summary>
    public WordImage AddImage(Uri uri, OfficeLength width, OfficeLength height, WordImageOptions? options = null) {
        var resolved = options ?? new WordImageOptions();
        if (resolved.Vml) throw new ArgumentException("VML images must be embedded from a file or stream.", nameof(options));
        return AttachImage(new WordImage(_document, this, uri, width.ToPixels(), height.ToPixels(), resolved.Wrapping, resolved.Description), resolved);
    }

    private WordImage AttachImage(WordImage image, WordImageOptions options) {
        VerifyRun().Append(image._Image);
        return options.Apply(image);
    }

    internal WordParagraph AddTrustedExternalImage(Uri uri, double width, double height,
        WordImageTextWrapping wrapping = WordImageTextWrapping.InLineWithText, string description = "") {
        if (uri == null) throw new ArgumentNullException(nameof(uri));
        if (!uri.IsFile) throw new ArgumentException("Trusted local external images must use the file scheme.", nameof(uri));
        var image = new WordImage(_document, this, uri, width, height, wrapping, description, shape: null, compressionQuality: null, allowFileUri: true);
        VerifyRun().Append(image._Image);
        return this;
    }
}
