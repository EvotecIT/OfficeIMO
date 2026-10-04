using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Mobile;

/// <summary>A real, editable document for exploring the reader without importing a file.</summary>
internal static class MobileSampleDocument {
    internal static void Save(string path) {
        var ink = PdfColor.FromRgb(28, 43, 64);
        var blue = PdfColor.FromRgb(47, 99, 233);
        var options = new PdfOptions {
            DefaultFont = PdfStandardFont.Helvetica,
            DefaultFontSize = 14,
            DefaultTextColor = ink,
            MarginLeft = 64, MarginRight = 64, MarginTop = 64, MarginBottom = 64,
            DefaultParagraphStyle = new PdfParagraphStyle { LineHeight = 1.5, SpacingAfter = 16 }
        };
        options.SetDefaultHeadingStyle(1, new PdfHeadingStyle { FontSize = 40, LineHeight = 1.1, Color = ink, SpacingAfter = 24 });
        options.SetDefaultHeadingStyle(2, new PdfHeadingStyle { FontSize = 20, Color = ink, SpacingBefore = 20, SpacingAfter = 12 });
        var document = PdfDocument.Create(pdf => pdf.Content(content => content
            .Paragraph(p => p.Text("OFFICEIMO STUDIO  /  A QUICK TOUR"), defaultColor: blue)
            .Spacer(58)
            .H1("A little room\nto think.")
            .Paragraph(p => p.Text("Your documents. Your thoughts.\nA workspace that keeps them together."))
            .Spacer(30)
            .H2("01   Read with room to breathe")
            .Paragraph(p => p.Text("Move between pages using the thumbnails. Pinch to look closer, or fit the whole page when you want the bigger picture."))
            .H2("02   Keep your place")
            .Paragraph(p => p.Text("Open another PDF in a tab. Your reading position and edits stay with each document when you switch."))
            .PageBreak()
            .Paragraph(p => p.Text("THE QUICK TOUR  /  REVIEW"), defaultColor: blue)
            .Spacer(32)
            .H1("Make a note.\nKeep the thought.")
            .Paragraph(p => p.Text("Good ideas rarely arrive in order. Leave a note on a page and come back to it when the time is right."))
            .H2("Find the detail")
            .Paragraph(p => p.Text("Search looks through the entire document. Matches stay highlighted on the page while you move through the results."))
            .H2("Change your mind")
            .Paragraph(p => p.Text("Undo and redo let you explore an edit. Studio keeps a recovery copy as you work."))
            .PageBreak()
            .Paragraph(p => p.Text("THE QUICK TOUR  /  SHARE"), defaultColor: blue)
            .Spacer(32)
            .H1("Ready when\nyou are.")
            .Paragraph(p => p.Text("Bring in a PDF from Files. Studio makes a working copy, so your original stays where you left it."))
            .H2("Take your work with you")
            .Paragraph(p => p.Text("Share saves the current document and opens the Apple share sheet. Save it to Files or choose where it goes next."))
            .H2("Pick up where you left off")
            .Paragraph(p => p.Text("Your open tabs and reading positions return when you reopen Studio. Unsaved review notes are recovered with their document."))), options);
        document.Meta(title: "Welcome to Studio", author: "OfficeIMO").Save(path);
    }
}
