using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects public column, inset and vertical-alignment authoring into the native creation specimen.</summary>
internal static class CandidateFrameLayout {
    internal static CandidatePublicationSpec Create(bool edited) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        document.AddPage();
        var layouts = new[] {
            (Columns: 1, Gap: 0D, Padding: new OfficeTextPadding(12, 6, 18, 9), Alignment: OfficeTextVerticalAlignment.Top),
            (Columns: 2, Gap: 12D, Padding: OfficeTextPadding.Empty, Alignment: OfficeTextVerticalAlignment.Top),
            (Columns: 3, Gap: 18D, Padding: OfficeTextPadding.Empty, Alignment: OfficeTextVerticalAlignment.Top),
            (Columns: 1, Gap: 0D, Padding: new OfficeTextPadding(12, 6, 18, 9), Alignment: OfficeTextVerticalAlignment.Center),
            (Columns: 1, Gap: 0D, Padding: OfficeTextPadding.Empty, Alignment: OfficeTextVerticalAlignment.Bottom),
            (Columns: 3, Gap: 12.1, Padding: new OfficeTextPadding(1.21, 2.42, 3.63, 4.84), Alignment: OfficeTextVerticalAlignment.Top)
        };
        for (int index = 0; index < layouts.Length; index++) {
            var layout = layouts[index];
            string text = index is 3 or 4 ? "Short native body." : string.Concat(Enumerable.Repeat("One two three four five six seven eight nine ten. ", 12)).TrimEnd();
            PublisherTextFrame frame = document.AddTextFrame(document.Pages[index / 3].Id,
                54, 54 + index % 3 * 228, 504, 180, new[] { new OfficeRichTextParagraph(new[] {
                    new OfficeRichTextRun(text, 12, OfficeColor.Black, fontFamily: "Arial")
                }, lineHeight: 18) });
            document.SetTextFrameLayout(frame.Id, edited ? 2 : layout.Columns, edited ? 14 : layout.Gap,
                edited ? new OfficeTextPadding(9, 3, 15, 6) : layout.Padding,
                edited ? OfficeTextVerticalAlignment.Bottom : layout.Alignment);
        }
        CandidateObjectSpec[] objects = document.Pages.SelectMany((page, pageIndex) => page.TextFrames.Select(frame =>
            new CandidateObjectSpec(CandidateObjectKind.TextBox, new CandidatePlacement(frame.X, frame.Y, frame.Width, frame.Height),
                CandidateStorySpec.Styled(document.TextStories.Single(story => story.Id == frame.StoryId).Paragraphs.Single()),
                PageIndex: pageIndex, FrameLayout: frame))).ToArray();
        return new CandidatePublicationSpec(612 * 12700, 792 * 12700, objects, document.Pages.Count);
    }
}
