using System;
using System.IO;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public class OpenDocumentSchemaArtifactTests {
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
                text.Save(Path.Combine(output, "schema-proof-1.4.odt"));
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
                spreadsheet.Save(Path.Combine(output, "schema-proof-1.3.ods"), new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.Odf13 });
                Assert.True(spreadsheet.Validate().IsValid);
            }
            using (OdpPresentation presentation = OdpPresentation.Create()) {
                OdpSlide slide = presentation.AddSlide("Schema proof");
                slide.AddTextBox(OdfRect.FromCentimeters(1, 1, 12, 2), "Native ODP");
                slide.AddRectangle(OdfRect.FromCentimeters(1, 4, 4, 2)).FillColor = OdfColor.Parse("#D1E9FF");
                slide.AddTable(OdfRect.FromCentimeters(7, 4, 8, 3), 2, 2, "Proof").Cell(0, 0).Text = "Value";
                slide.GetOrCreateSpeakerNotes().AddParagraph("Speaker notes");
                slide.TransitionType = "automatic";
                slide.TransitionStyle = "fade-from-center";
                presentation.Save(Path.Combine(output, "schema-proof-1.4.odp"));
                presentation.Save(Path.Combine(output, "schema-proof-1.3.odp"), new OdfSaveOptions { CompatibilityProfile = OdfCompatibilityProfile.Odf13 });
                Assert.True(presentation.Validate().IsValid);
            }
        } finally {
            if (!keep && Directory.Exists(output)) Directory.Delete(output, recursive: true);
        }
    }
}
