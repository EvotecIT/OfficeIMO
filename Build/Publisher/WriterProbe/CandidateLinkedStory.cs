using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects a public-model chain whose order differs from physical page order.</summary>
internal static class CandidateLinkedStory {
    internal static CandidatePublicationSpec Create(bool overflow, uint? languageId = null) =>
        CandidateModelProjection.LinkedStories(Author(overflow, languageId));

    internal static PublisherDocument Author(bool overflow, uint? languageId = null) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        document.AddPage(master.Id);
        document.AddPage(master.Id);
        string text = string.Concat(Enumerable.Repeat(
            "Linked stories preserve reading order across publication pages. ", overflow ? 24 : 6)).TrimEnd();
        PublisherTextFrame first = document.AddTextFrame(document.Pages[0].Id, 72, 72, 216, 72, Paragraph(text, 12, languageId));
        document.AddTextFrame(document.Pages[0].Id, 360, 72, 180, 54, Paragraph("Independent story", 14, languageId));
        PublisherTextFrame middle = document.AddLinkedTextFrame(first.Id, document.Pages[2].Id, 72, 144, 216, 72);
        document.AddLinkedTextFrame(middle.Id, document.Pages[1].Id, 72, 216, 360, 108);
        return document;
    }

    private static OfficeRichTextParagraph[] Paragraph(string text, double size, uint? languageId) => new[] {
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, OfficeColor.Black, fontFamily: "Arial").WithLanguageId(languageId) }, lineHeight: 18)
    };
}
