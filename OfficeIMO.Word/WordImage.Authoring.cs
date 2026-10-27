using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordImage {
    /// <summary>Gets the paragraph containing this image occurrence.</summary>
    public WordParagraph Paragraph {
        get {
            var paragraph = OccurrenceElement.Ancestors<Paragraph>().FirstOrDefault()
                ?? throw new InvalidOperationException("The image is not attached to a paragraph.");
            var run = OccurrenceElement.Ancestors<Run>().FirstOrDefault();
            return run == null ? new WordParagraph(_document, paragraph) : new WordParagraph(_document, paragraph, run);
        }
    }

    /// <summary>Gets or sets the external hyperlink wrapping this image. Null removes the link.</summary>
    public Uri? Hyperlink {
        get {
            var link = OccurrenceElement.Ancestors<Hyperlink>().FirstOrDefault();
            return link == null ? null : new WordHyperLink(_document, Paragraph._paragraph, link).Uri;
        }
        set {
            var run = OccurrenceElement.Ancestors<Run>().FirstOrDefault()
                ?? throw new InvalidOperationException("The image is not attached to a run.");
            var paragraph = Paragraph;
            if (run.Parent is Hyperlink existing) {
                if (value != null) new WordHyperLink(_document, paragraph._paragraph, existing).Uri = value;
                else {
                    foreach (var content in existing.ChildElements.ToArray()) {
                        content.Remove();
                        existing.InsertBeforeSelf(content);
                    }
                    existing.Remove();
                }
                return;
            }
            if (value == null) return;
            run = IsolateImageRun(run);
            var owner = WordHyperLink.ResolveRelationshipPart(_document, paragraph._paragraph)
                ?? throw new InvalidOperationException("The image's relationship owner is missing.");
            var relationship = owner.AddHyperlinkRelationship(value, true);
            var hyperlink = new Hyperlink { Id = relationship.Id };
            run.InsertBeforeSelf(hyperlink);
            run.Remove();
            hyperlink.Append(run);
        }
    }

    private Run IsolateImageRun(Run run) {
        OpenXmlElement content = OccurrenceElement;
        while (content.Parent != run) content = content.Parent
            ?? throw new InvalidOperationException("The image is not contained in its run.");
        var siblings = run.ChildElements.Where(element => element is not RunProperties).ToArray();
        if (siblings.Length == 1) return run;
        var isolated = new Run();
        if (run.RunProperties != null) isolated.RunProperties = (RunProperties)run.RunProperties.CloneNode(true);
        Run? suffix = null;
        foreach (var following in siblings.SkipWhile(element => element != content).Skip(1)) {
            if (suffix == null) {
                suffix = new Run();
                if (run.RunProperties != null) suffix.RunProperties = (RunProperties)run.RunProperties.CloneNode(true);
                run.InsertAfterSelf(suffix);
            }
            following.Remove();
            suffix.Append(following);
        }
        run.InsertAfterSelf(isolated);
        content.Remove();
        isolated.Append(content);
        return isolated;
    }

    /// <summary>Clamps the displayed width while retaining the image's other authoring properties.</summary>
    public WordImage SetMaximumWidth(OfficeLength width) {
        if (width <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(width));
        if (Width is not OfficeLength current || current > width) Width = width;
        return this;
    }
}
