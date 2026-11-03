using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Mobile;

/// <summary>A real, editable document for exploring the reader without importing a file.</summary>
internal static class MobileSampleDocument {
    internal static void Save(string path) {
        var ink = global::OfficeIMO.OfficeColor.FromRgb(28, 43, 64);
        var blue = global::OfficeIMO.OfficeColor.FromRgb(47, 99, 233);
        var options = new PdfCreateOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = global::OfficeIMO.OfficeLength.Points(14),
            DefaultTextColor = ink,
            PageSetup = {
                Margins = { Left = global::OfficeIMO.OfficeLength.Points(64), Right = global::OfficeIMO.OfficeLength.Points(64), Top = global::OfficeIMO.OfficeLength.Points(64), Bottom = global::OfficeIMO.OfficeLength.Points(64) }
            },
            DefaultParagraphStyle = new PdfParagraphStyle { LineHeight = 1.5, SpacingAfter = global::OfficeIMO.OfficeLength.Points(16) }
        };
        options.SetDefaultHeadingStyle(1, new PdfHeadingStyle { FontSize = global::OfficeIMO.OfficeLength.Points(40), LineHeight = 1.1, Color = ink, SpacingAfter = global::OfficeIMO.OfficeLength.Points(24) });
        options.SetDefaultHeadingStyle(2, new PdfHeadingStyle { FontSize = global::OfficeIMO.OfficeLength.Points(20), Color = ink, SpacingBefore = global::OfficeIMO.OfficeLength.Points(20), SpacingAfter = global::OfficeIMO.OfficeLength.Points(12) });
        var document = PdfDocument.Create(pdf => pdf.Content(content => content
            .AddParagraph(p => p.AddText("OFFICEIMO STUDIO  /  A QUICK TOUR"), color: blue)
            .Spacer(global::OfficeIMO.OfficeLength.Points(58))
            .H1("A little room\nto think.")
            .AddParagraph(p => p.AddText("Your documents. Your thoughts.\nA workspace that keeps them together."))
            .Spacer(global::OfficeIMO.OfficeLength.Points(30))
            .H2("01   Read with room to breathe")
            .AddParagraph(p => p.AddText("Move between pages using the thumbnails. Pinch to look closer, or fit the whole page when you want the bigger picture."))
            .H2("02   Keep your place")
            .AddParagraph(p => p.AddText("Open another PDF in a tab. Your reading position and edits stay with each document when you switch."))
            .PageBreak()
            .AddParagraph(p => p.AddText("THE QUICK TOUR  /  REVIEW"), color: blue)
            .Spacer(global::OfficeIMO.OfficeLength.Points(32))
            .H1("Make a note.\nKeep the thought.")
            .AddParagraph(p => p.AddText("Good ideas rarely arrive in order. Leave a note on a page and come back to it when the time is right."))
            .H2("Find the detail")
            .AddParagraph(p => p.AddText("Search looks through the entire document. Matches stay highlighted on the page while you move through the results."))
            .H2("Change your mind")
            .AddParagraph(p => p.AddText("Undo and redo let you explore an edit. Studio keeps a recovery copy as you work."))
            .PageBreak()
            .AddParagraph(p => p.AddText("THE QUICK TOUR  /  SHARE"), color: blue)
            .Spacer(global::OfficeIMO.OfficeLength.Points(32))
            .H1("Ready when\nyou are.")
            .AddParagraph(p => p.AddText("Bring in a PDF from Files. Studio makes a working copy, so your original stays where you left it."))
            .H2("Take your work with you")
            .AddParagraph(p => p.AddText("Share saves the current document and opens the Apple share sheet. Save it to Files or choose where it goes next."))
            .H2("Pick up where you left off")
            .AddParagraph(p => p.AddText("Your open tabs and reading positions return when you reopen Studio. Unsaved review notes are recovered with their document."))), options);
        document.Meta(title: "Welcome to Studio", author: "OfficeIMO").Save(path);
    }
}
