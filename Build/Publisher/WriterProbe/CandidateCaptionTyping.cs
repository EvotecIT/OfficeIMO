using OfficeIMO.Drawing;
using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher.WriterProbe;

/// <summary>Projects one ordinary caption with inherited last-run typing style and optional explicit language.</summary>
internal static class CandidateCaptionTyping {
    internal static PublisherNativeCreationPlan Create(string? text, uint? languageId) {
        PublisherDocument document = PublisherDocument.Create(612, 792);
        PublisherPage master = document.AddMasterPage("Master Page A");
        uint page = document.Pages[0].Id;
        document.SetPageMaster(page, master.Id);
        document.AddTextFrame(page, 72, 72, 300, 54, new[] {
            new OfficeRichTextParagraph(new[] {
                new OfficeRichTextRun(text ?? "Independent table label", 12, OfficeColor.Black, fontFamily: "Arial")
                    .WithLanguageId(languageId)
            }, lineHeight: 18)
        });
        return PublisherNativeCreationProjection.FrameLayouts(document);
    }
}
