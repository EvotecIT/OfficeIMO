using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Authors ordinary paragraph and CR-typing controls, retaining complete overflow in one frame or a reordered page chain.</summary>
internal static class CandidateMultiParagraph {
    internal static PublisherNativeCreationPlan Create(bool linked, bool automaticHyphenation = false) {
        PublisherDocument document = Author(linked, automaticHyphenation);
        return linked ? PublisherNativeCreationProjection.LinkedStories(document)
            : PublisherNativeCreationProjection.FrameLayouts(document);
    }

    internal static PublisherDocument Author(bool linked, bool automaticHyphenation = false) {
        PublisherDocument document = PublisherDocument.Create(OfficeLength.Points(612), OfficeLength.Points(792));
        PublisherPage master = document.AddMasterPage("Master Page A");
        document.SetPageMaster(document.Pages[0].Id, master.Id);
        if (linked) {
            document.AddPage(master.Id);
            document.AddPage(master.Id);
        }
        PublisherTextFrame first = document.AddTextFrame(document.Pages[0].Id, OfficeLength.Points(72), OfficeLength.Points(72), OfficeLength.Points(216), OfficeLength.Points(72), Paragraphs());
        document.SetStoryAutomaticHyphenation(first.StoryId, automaticHyphenation);
        if (linked) {
            PublisherTextFrame middle = document.AddLinkedTextFrame(first.Id, document.Pages[2].Id, OfficeLength.Points(72), OfficeLength.Points(144), OfficeLength.Points(216), OfficeLength.Points(72));
            document.AddLinkedTextFrame(middle.Id, document.Pages[1].Id, OfficeLength.Points(72), OfficeLength.Points(216), OfficeLength.Points(360), OfficeLength.Points(108));
        }
        return document;
    }

    internal static OfficeRichTextParagraph[] Paragraphs() => new[] {
        new OfficeRichTextParagraph(new[] {
            Run("Multi-paragraph heading. ", 16, true), Run("Body follows below.", 12, false)
        }, lineHeight: 22),
        new OfficeRichTextParagraph(new[] { Run(string.Empty, 14, true) }, lineHeight: 20),
        new OfficeRichTextParagraph(new[] {
            Run("Final paragraph: ", 12, true),
            Run(string.Concat(Enumerable.Repeat(
                "Complete ordinary stories retain every paragraph and all overflow text. ", 18)).TrimEnd(), 12, false)
        }, lineHeight: 18)
    };

    private static OfficeRichTextRun Run(string text, double size, bool bold) =>
        new OfficeRichTextRun(text, size, OfficeColor.Black, bold, fontFamily: "Arial").WithLanguageId(1033);
}
