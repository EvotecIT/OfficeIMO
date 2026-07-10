using System;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentSchemaArtifactTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    [Trait("Category", "OpenDocumentSchemaArtifact")]
    public void EmitsRepresentativeOdf14Artifacts() {
        string? requestedOutput = Environment.GetEnvironmentVariable("OFFICEIMO_ODF_SCHEMA_OUTPUT");
        bool keep = !string.IsNullOrWhiteSpace(requestedOutput);
        string output = keep ? Path.GetFullPath(requestedOutput!) : Path.Combine(Path.GetTempPath(), "OfficeIMO-ODF-Schema-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(output);
        try {
            using (OdtDocument text = OdtDocument.Create()) {
                text.AddHeading("Schema proof", 1);
                text.AddParagraph("Native ODT").AddSpan(" with formatting").Bold = true;
                text.AddList().AddItem("One");
                text.AddTable(2, 2, "Proof").Cell(0, 0).Text = "Value";
                text.PageLayout.Header.AddParagraph("OfficeIMO");
                text.AddTrackedParagraphInsertion("Tracked schema proof", "OfficeIMO", new DateTimeOffset(2026, 7, 10, 0, 0, 0, TimeSpan.Zero));
                text.AddParagraph("Embedded image").AddImage(TinyPng, "pixel.png", OdfLength.Centimeters(1), OdfLength.Centimeters(1));
                text.Save(Path.Combine(output, "schema-proof-1.4.odt"));
                text.SaveFlatXml(Path.Combine(output, "schema-proof-1.4.fodt"));
                text.Save(Path.Combine(output, "schema-proof-1.3.odt"), new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.Odf13 });
                Assert.True(text.Validate().IsValid);
            }
            using (OdsDocument spreadsheet = OdsDocument.Create()) {
                OdsSheet sheet = spreadsheet.AddSheet("Data");
                sheet.Cell(0, 0).SetString("Value");
                OdsCell formula = sheet.Cell(1, 0);
                formula.Formula = "of:=SUM([.A1:.A1])";
                formula.SetDecimal(1m);
                formula.NumberFormatName = spreadsheet.AddNumberStyle("Amount", 2).Name;
                spreadsheet.Save(Path.Combine(output, "schema-proof-1.4.ods"));
                spreadsheet.SaveFlatXml(Path.Combine(output, "schema-proof-1.4.fods"));
                spreadsheet.Save(Path.Combine(output, "schema-proof-1.3.ods"), new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.Odf13 });
                Assert.True(spreadsheet.Validate().IsValid);
            }
            using (OdpPresentation presentation = OdpPresentation.Create()) {
                OdpSlide slide = presentation.AddSlide("Schema proof");
                slide.AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 2), "Native ODP");
                OdpRectangle rectangle = slide.AddRectangle(OdfRect.FromCentimeters(1, 4, 4, 2));
                rectangle.FillColor = OdfColor.Parse("#D1E9FF");
                slide.AddFadeInAnimation(rectangle, TimeSpan.FromSeconds(1));
                slide.AddTable(OdfRect.FromCentimeters(7, 4, 8, 3), 2, 2, "Proof").Cell(0, 0).Text = "Value";
                slide.GetOrCreateSpeakerNotes().AddParagraph("Speaker notes");
                slide.TransitionType = "automatic";
                slide.TransitionStyle = "fade-from-center";
                presentation.Save(Path.Combine(output, "schema-proof-1.4.odp"));
                presentation.SaveFlatXml(Path.Combine(output, "schema-proof-1.4.fodp"));
                presentation.Save(Path.Combine(output, "schema-proof-1.3.odp"), new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.Odf13 });
                Assert.True(presentation.Validate().IsValid);
            }
        } finally {
            if (!keep && Directory.Exists(output)) Directory.Delete(output, recursive: true);
        }
    }

    [Fact]
    [Trait("Category", "OpenDocumentLibreOfficeArtifact")]
    public void ReopensLibreOfficeResavedArtifactsWithExpectedSemantics() {
        string? requestedInput = Environment.GetEnvironmentVariable("OFFICEIMO_ODF_INTEROP_INPUT");
        if (string.IsNullOrWhiteSpace(requestedInput)) return;
        string input = Path.GetFullPath(requestedInput!);
        string[] files = Directory.GetFiles(input, "*.*", SearchOption.AllDirectories)
            .Where(path => path.EndsWith(".odt", StringComparison.OrdinalIgnoreCase) ||
                path.EndsWith(".ods", StringComparison.OrdinalIgnoreCase) ||
                path.EndsWith(".odp", StringComparison.OrdinalIgnoreCase)).ToArray();
        Assert.Equal(6, files.Length);

        foreach (string path in files) {
            using OdfDocument document = OdfDocument.OpenAny(path);
            OdfValidationResult validation = document.Validate();
            Assert.True(validation.IsValid, string.Join(Environment.NewLine, validation.Diagnostics.Select(item => item.Id + ": " + item.Message)));
            if (document is OdtDocument text) {
                Assert.Contains(text.ContentBlocks, block => block.Paragraph?.Text.Contains("Schema proof", StringComparison.Ordinal) == true);
            } else if (document is OdsDocument spreadsheet) {
                Assert.Equal("Value", spreadsheet.GetSheet("Data")!.GetValue(0, 0).DisplayText);
            } else if (document is OdpPresentation presentation) {
                Assert.Contains(presentation.Slides.SelectMany(slide => slide.Shapes).OfType<OdpTextBox>(),
                    box => box.Paragraphs.Any(paragraph => paragraph.Text.Contains("Native ODP", StringComparison.Ordinal)));
            }
        }
    }
}
