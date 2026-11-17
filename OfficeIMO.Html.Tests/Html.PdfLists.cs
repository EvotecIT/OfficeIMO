using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfListTests {
    [Fact]
    public void Pdf_ToHtml_NestsListsAndKeepsOrderedListsApart() {
        var flat = new PdfCore.PdfListStyle { SpacingBefore = global::OfficeIMO.OfficeLength.Zero };
        var nested = new PdfCore.PdfListStyle { LeftIndent = global::OfficeIMO.OfficeLength.Points(18), SpacingBefore = global::OfficeIMO.OfficeLength.Zero };
        var document = PdfCore.PdfDocument.Create();
        document.AddList(new[] { "Reduce backlog" }, style: flat);
        document.AddList(new[] { "Archive exports", "Notify owners" }, style: nested);
        document.AddList(new[] { "Expand coverage" }, style: flat);
        document.AddParagraph("Next steps follow.");
        document.AddNumberedList(new[] { "Collect requirements", "Agree on scope" });
        document.AddParagraph("Follow-up actions.");
        document.AddNumberedList(new[] { "Add alert", "Rehearse" }, startNumber: 3);

        string html = PdfCore.PdfDocumentReadResult.Load(document.ToBytes(), new PdfCore.PdfTextLayoutOptions { ForceSingleColumn = true })
            .ToHtml(PdfHtmlExportOptions.CreateSemanticProfile());

        Assert.Contains("<ul><li>Reduce backlog<ul><li>Archive exports</li><li>Notify owners</li></ul></li><li>Expand coverage</li></ul>", html, StringComparison.Ordinal);
        Assert.Contains("<ol><li>Collect requirements</li><li>Agree on scope</li></ol>", html, StringComparison.Ordinal);
        Assert.Contains("<ol start=\"3\"><li>Add alert</li><li>Rehearse</li></ol>", html, StringComparison.Ordinal);
    }
}
