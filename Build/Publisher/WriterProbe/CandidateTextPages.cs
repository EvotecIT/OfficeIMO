using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Exercises the existing public authoring model with independent stories on the first and last physical pages.</summary>
internal static class CandidateTextPages {
    internal static CandidatePublicationSpec Create(string? value) {
        if (!int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int count)
            || count < 2 || count > 16) throw new ArgumentException("The text-pages case requires a page count from 2 to 16.");
        PublisherDocument document = PublisherDocument.Create(612, 792);
        document.AddTextFrame(document.Pages[0].Id, 72, 72, 288, 108,
            Paragraph("First page: native authoring", 12));
        for (int index = 1; index < count; index++) document.AddPage();
        document.AddTextFrame(document.Pages[count - 1].Id, 108, 144, 360, 90,
            Paragraph("Last page: native authoring", 18));

        // This closed specimen contains only the two independent text frames.
        // Read their current model views instead of reconstructing story content
        // or page placement separately in the format codecs.
        CandidateObjectSpec[] objects = document.Pages.SelectMany((page, pageIndex) => page.TextFrames.Select(frame =>
            new CandidateObjectSpec(CandidateObjectKind.TextBox,
                new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height),
                CandidateStorySpec.Styled(document.TextStories.Single(story => story.Id == frame.StoryId).Paragraphs.Single()),
                PageIndex: pageIndex, ColumnSpacing: frame.ColumnSpacing))).ToArray();
        return new CandidatePublicationSpec(checked((uint)(document.Pages[0].Width * 12700)),
            checked((uint)(document.Pages[0].Height * 12700)), objects, document.Pages.Count);
    }

    private static OfficeRichTextParagraph[] Paragraph(string text, double size) => new[] {
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, OfficeColor.Black, fontFamily: "Arial") })
    };
}
