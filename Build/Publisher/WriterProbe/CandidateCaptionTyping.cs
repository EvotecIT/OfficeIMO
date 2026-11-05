using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects one ordinary caption with inherited last-run typing style and optional explicit language.</summary>
internal static class CandidateCaptionTyping {
    internal static PublisherNativeCreationPlan Create(string? text, uint? languageId) {
        PublisherDocument document = PublisherDocument.Create(OfficeLength.Points(612), OfficeLength.Points(792));
        PublisherPage master = document.AddMasterPage("Master Page A");
        uint page = document.Pages[0].Id;
        document.SetPageMaster(page, master.Id);
        document.AddTextFrame(page, OfficeLength.Points(72), OfficeLength.Points(72), OfficeLength.Points(300), OfficeLength.Points(54), new[] {
            new OfficeRichTextParagraph(new[] {
                new OfficeRichTextRun(text ?? "Independent table label", 12, OfficeColor.Black, fontFamily: "Arial")
                    .WithLanguageId(languageId)
            }, lineHeight: 18)
        });
        return PublisherNativeCreationProjection.FrameLayouts(document);
    }
}
