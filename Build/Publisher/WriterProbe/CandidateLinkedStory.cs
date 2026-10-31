using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects a public-model chain whose order differs from physical page order.</summary>
internal static class CandidateLinkedStory {
    internal static CandidatePublicationSpec Create(bool overflow) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        document.AddPage();
        document.AddPage();
        string text = string.Concat(Enumerable.Repeat(
            "Linked stories preserve reading order across publication pages. ", overflow ? 24 : 6)).TrimEnd();
        PublisherTextFrame first = document.AddTextFrame(document.Pages[0].Id, 72, 72, 216, 72, Paragraph(text, 12));
        document.AddTextFrame(document.Pages[0].Id, 360, 72, 180, 54, Paragraph("Independent story", 14));
        PublisherTextFrame middle = document.AddLinkedTextFrame(first.Id, document.Pages[2].Id, 72, 144, 216, 72);
        document.AddLinkedTextFrame(middle.Id, document.Pages[1].Id, 72, 216, 360, 108);
        // One story object per model identity keeps Quill text and formatting
        // independent of how many frames reference it or how pages are ordered.
        var stories = document.TextStories.ToDictionary(story => story.Id,
            story => CandidateStorySpec.Styled(story.Paragraphs.Single()));
        CandidateObjectSpec[] objects = document.Pages.SelectMany((page, pageIndex) =>
            page.TextFrames.Select(frame => (Frame: frame, PageIndex: pageIndex)))
            .OrderBy(item => item.Frame.Id).Select(item => {
                PublisherTextFrame frame = item.Frame;
                return new CandidateObjectSpec(CandidateObjectKind.TextBox,
                new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height), stories[frame.StoryId],
                PageIndex: item.PageIndex, ColumnSpacing: frame.ColumnSpacing, FrameOrder: frame.Order);
            }).ToArray();
        return new CandidatePublicationSpec(612 * 12700, 792 * 12700, objects, document.Pages.Count);
    }

    private static OfficeRichTextParagraph[] Paragraph(string text, double size) => new[] {
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, OfficeColor.Black, fontFamily: "Arial") }, lineHeight: 18)
    };
}
