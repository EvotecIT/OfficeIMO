using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Authors the independent multiple-master controls through the canonical public page and story model.</summary>
internal static class CandidateMasterPages {
    internal static PublisherNativeCreationPlan Create(bool paired) => PublisherNativeCreationProjection.MultipleMasters(Author(paired));

    internal static PublisherDocument Author(bool paired) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage centered = document.AddMasterPage("Master Page A");
        PublisherPage[] masters;
        int[] assignments;
        if (paired) {
            var pair = document.AddMasterPagePair("Comparison Two-Page Master");
            masters = new[] { centered, pair.Left, pair.Right };
            assignments = new[] { 0, 1, 2 };
        } else {
            masters = new[] { centered, document.AddMasterPage("Comparison Master B") };
            assignments = new[] { 0, 1, 0 };
        }
        document.SetPageMaster(document.Pages[0].Id, masters[assignments[0]].Id);
        document.AddPage(masters[assignments[1]].Id);
        document.AddPage(masters[assignments[2]].Id);
        for (int index = 0; index < masters.Length; index++) {
            string text = paired ? new[] { "Default master", "Paired master left", "Paired master right" }[index]
                : index == 0 ? "Master Alpha" : "Master Beta";
            Text(masters[index].Id, text, paired ? index == 2 ? 342 : 18 : index == 1 ? 342 : 18, 24, 252, 24, 10);
        }
        for (int index = 0; index < document.Pages.Count; index++)
            Text(document.Pages[index].Id, (paired ? "TwoPageMaster" : "MultipleMasters") + " page " + (index + 1), 72, 96, 468, 48, 16);
        return document;

        void Text(uint owner, string text, double x, double y, double width, double height, double size) {
            document.AddTextFrame(owner, x, y, width, height, new[] {
                new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(text, size, OfficeColor.Black, fontFamily: "Arial") })
            });
        }
    }
}
