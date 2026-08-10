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
            {
                OdtDocument text = OdtDocument.Create();
                text.AddHeading("Schema proof", 1);
                OdtParagraph richText = text.AddParagraph("Native ODT ");
                richText.Alignment = OdtParagraphAlignment.Center;
                richText.IndentStart = OdfLength.Points(12);
                OdtSpan richSpan = richText.AddSpan("with formatting");
                richSpan.Bold = true;
                richSpan.Underline = true;
                richSpan.BackgroundColor = OdfColor.Parse("#FFF200");
                richText.AddText(" and ");
                richText.AddHyperlink("a link", "https://example.com").Italic = true;
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
            {
                OdsDocument spreadsheet = OdsDocument.Create();
                OdsSheet sheet = spreadsheet.AddSheet("Data");
                sheet.Cell(0, 0).SetString("Value");
                OdsCell formula = sheet.Cell(1, 0);
                formula.Formula = "of:=SUM([.A1:.A1])";
                formula.SetDecimal(1m);
                formula.NumberFormatName = spreadsheet.AddNumberStyle("Amount", 2).Name;
                formula.AddAnnotation("Calculated value", "OfficeIMO");
                OdsValidation validation = spreadsheet.AddValidation("PositiveWholeNumber",
                    OdsValidationConditionSyntax.Create(OdsValidationValueKind.WholeNumber,
                        OdsValidationComparison.GreaterThan, "0"));
                validation.SetHelpMessage("Input", "Enter a positive whole number.");
                formula.ValidationName = validation.Name;
                spreadsheet.Save(Path.Combine(output, "schema-proof-1.4.ods"));
                spreadsheet.SaveFlatXml(Path.Combine(output, "schema-proof-1.4.fods"));
                spreadsheet.Save(Path.Combine(output, "schema-proof-1.3.ods"), new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.Odf13 });
                Assert.True(spreadsheet.Validate().IsValid);
            }
            {
                OdpPresentation presentation = OdpPresentation.Create();
                OdpSlide slide = presentation.AddSlide("Schema proof");
                OdpParagraph presentationText = slide.AddTextBox(
                    OdfRect.FromCentimeters(1, 1, 12, 2), null).AddParagraph();
                presentationText.AddText("Native ODP ");
                OdpRun presentationRun = presentationText.AddRun("with formatting");
                presentationRun.Bold = true;
                presentationRun.StrikeThrough = true;
                presentationRun.BackgroundColor = OdfColor.Parse("#FFF200");
                presentationText.AddText(" and ");
                presentationText.AddHyperlink("a link", "https://example.com").Underline = true;
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
            OdfDocument document = OdfDocument.Load(path);
            OdfValidationResult validation = document.Validate();
            Assert.True(validation.IsValid, string.Join(Environment.NewLine, validation.Diagnostics.Select(item => item.Id + ": " + item.Message)));
            if (document is OdtDocument text) {
                Assert.Contains(text.ContentBlocks, block => block.Paragraph?.Text.IndexOf("Schema proof", StringComparison.Ordinal) >= 0);
                OdtParagraph rich = text.Paragraphs.Single(paragraph => paragraph.Text.IndexOf("Native ODT", StringComparison.Ordinal) >= 0);
                Assert.Contains(rich.InlineNodes, node => node.Kind == OdtInlineNodeKind.Span && node.Span!.Underline == true);
                Assert.Contains(rich.InlineNodes, node => node.Kind == OdtInlineNodeKind.Hyperlink &&
                    Uri.Compare(new Uri(node.Hyperlink!.Href), new Uri("https://example.com"),
                        UriComponents.AbsoluteUri, UriFormat.SafeUnescaped, StringComparison.OrdinalIgnoreCase) == 0);
            } else if (document is OdsDocument spreadsheet) {
                OdsSheet sheet = spreadsheet.GetSheet("Data")!;
                Assert.Equal("Value", sheet.GetValue(0, 0).DisplayText);
                OdsCell formula = sheet.Cell(1, 0);
                Assert.Contains(formula.Annotations, annotation => annotation.Text == "Calculated value" && annotation.Creator == "OfficeIMO");
                Assert.False(string.IsNullOrWhiteSpace(formula.ValidationName));
                Assert.Contains(spreadsheet.Validations, item => item.ParsedCondition?.ValueKind == OdsValidationValueKind.WholeNumber);
            } else if (document is OdpPresentation presentation) {
                OdpParagraph rich = presentation.Slides.SelectMany(slide => slide.Shapes).OfType<OdpTextBox>()
                    .SelectMany(box => box.Paragraphs).Single(paragraph => paragraph.Text.IndexOf("Native ODP", StringComparison.Ordinal) >= 0);
                Assert.Contains(rich.InlineNodes, node => node.Kind == OdpInlineNodeKind.Run && node.Run!.StrikeThrough == true);
                Assert.Contains(rich.InlineNodes, node => node.Kind == OdpInlineNodeKind.Hyperlink &&
                    Uri.Compare(new Uri(node.Hyperlink!.Href), new Uri("https://example.com"),
                        UriComponents.AbsoluteUri, UriFormat.SafeUnescaped, StringComparison.OrdinalIgnoreCase) == 0);
            }
        }
    }
}
