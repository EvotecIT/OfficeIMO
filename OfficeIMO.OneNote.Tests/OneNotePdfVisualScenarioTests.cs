using OfficeIMO.OneNote.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.OneNote.Tests;

public sealed class OneNotePdfVisualScenarioTests {
    [Fact]
    public void OneNoteSemanticDocumentScenarioProducesArtifactAndExplicitLossWarnings() {
        var section = new OneNoteSection { Name = "OneNote semantic PDF proof" };
        var page = new OneNotePage { Title = "Positioned planning page", Width = 640, Height = 480 };
        var outline = new OneNoteOutline {
            Layout = new OneNoteLayout { X = 96, Y = 72, Width = 360, Height = 160 }
        };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Positioned note body" });
        outline.Children.Add(paragraph);
        page.Outlines.Add(outline);
        page.DirectContent.Add(new OneNoteImage {
            FileName = "diagram.png",
            AltText = "Unresolved planning diagram",
            MediaType = "image/png",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 })
        });
        section.Pages.Add(page);

        PdfCore.PdfDocumentConversionResult result = section.ToPdfDocumentResult();
        byte[] pdf = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("OneNote semantic PDF proof", text);
        Assert.Contains("Positioned note body", text);
        Assert.Contains(result.Warnings, warning => warning.Code == "ONENOTE_MARKDOWN_CANVAS_FLATTENED");
        Assert.Contains(result.Warnings, warning => warning.Code == "ONENOTE_MARKDOWN_ASSET_PLACEHOLDER");

        string? outputDirectory = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT");
        if (!string.IsNullOrWhiteSpace(outputDirectory)) {
            Directory.CreateDirectory(outputDirectory!);
            File.WriteAllBytes(Path.Combine(outputDirectory!, "onenote-semantic-document.pdf"), pdf);
        }
    }
}
