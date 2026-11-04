using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects public column, inset and vertical-alignment authoring into the native creation specimen.</summary>
internal static class CandidateFrameLayout {
    internal static PublisherNativeCreationPlan Create(bool edited, uint? languageId = null) =>
        PublisherNativeCreationProjection.FrameLayouts(Author(edited, languageId));

    internal static PublisherDocument Author(bool edited, uint? languageId = null) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        document.AddPage(master.Id);
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
                    new OfficeRichTextRun(text, 12, OfficeColor.Black, fontFamily: "Arial").WithLanguageId(languageId)
                }, lineHeight: 18) });
            document.SetTextFrameLayout(frame.Id, edited ? 2 : layout.Columns, edited ? 14 : layout.Gap,
                edited ? new OfficeTextPadding(9, 3, 15, 6) : layout.Padding,
                edited ? OfficeTextVerticalAlignment.Bottom : layout.Alignment);
        }
        return document;
    }
}
