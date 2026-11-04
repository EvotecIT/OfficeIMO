using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Exercises the existing public authoring model with independent stories on the first and last physical pages.</summary>
internal static class CandidateTextPages {
    internal static PublisherNativeCreationPlan Create(string? value) => PublisherNativeCreationProjection.SingleMasterContent(Author(value));

    internal static PublisherDocument Author(string? value) {
        if (!int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out int count)
            || count < 2 || count > 16) throw new ArgumentException("The text-pages case requires a page count from 2 to 16.");
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        document.AddTextFrame(document.Pages[0].Id, 72, 72, 288, 108,
            Paragraph("First page: native authoring", 12));
        for (int index = 1; index < count; index++) document.AddPage(master.Id);
        document.AddTextFrame(document.Pages[count - 1].Id, 108, 144, 360, 90,
            Paragraph("Last page: native authoring", 18));

        return document;
    }

    private static OfficeRichTextParagraph[] Paragraph(string text, double size) => new[] {
        new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, OfficeColor.Black, fontFamily: "Arial") })
    };
}
