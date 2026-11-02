namespace OfficeIMO.Word;

/// <summary>Controls image placement and metadata when adding an image to a Word document.</summary>
public sealed class WordImageOptions {
    /// <summary>Gets or sets the source name used for an image stream.</summary>
    public string FileName { get; set; } = "image.png";
    /// <summary>Gets or sets text wrapping.</summary>
    public WordImageTextWrapping Wrapping { get; set; } = WordImageTextWrapping.InLineWithText;
    /// <summary>Gets or sets the image title.</summary>
    public string? Title { get; set; }
    /// <summary>Gets or sets the alternative text description.</summary>
    public string Description { get; set; } = string.Empty;
    /// <summary>Gets or sets an external image hyperlink.</summary>
    public Uri? Hyperlink { get; set; }
    /// <summary>Gets or sets an optional displayed width limit.</summary>
    public OfficeLength? MaximumWidth { get; set; }
    /// <summary>Gets or sets the image paragraph's alignment.</summary>
    public WordParagraphAlignment? Alignment { get; set; }

    /// <summary>Gets or sets whether to embed legacy VML markup instead of DrawingML.</summary>
    public bool Vml { get; set; }

    internal WordImageOptions Clone() => (WordImageOptions)MemberwiseClone();

    internal WordImage Apply(WordImage image) {
        if (Title != null) image.Title = Title;
        if (Hyperlink != null) image.Hyperlink = Hyperlink;
        if (MaximumWidth.HasValue) image.SetMaximumWidth(MaximumWidth.Value);
        if (Alignment.HasValue) image.Paragraph.ParagraphAlignment = Alignment;
        return image;
    }
}
