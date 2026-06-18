using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointFeatureReportTests {
        [Fact]
        public void PowerPointFeatureReport_DetectsEditableAndPartiallyEditableFeatures() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointTextBox textBox = slide.AddTextBox("Quarterly review");
                    textBox.Paragraphs.First().Runs.First().SetHyperlink("https://example.com/review");
                    slide.AddPicture(imagePath);
                    PowerPointTable table = slide.AddTable(2, 2);
                    table.GetCell(0, 0).Text = "Metric";
                    slide.AddChart();
                    slide.Notes.Text = "Talk track";
                    slide.Transition = SlideTransition.Fade;
                    presentation.AddSection("Results", 0);
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Slides" && feature.Count == 1);
                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Text boxes" && feature.Count == 1);
                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Tables" && feature.Count == 1);
                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Table style metadata"
                        && feature.Details.Any(detail => detail.Contains("colIds=2", StringComparison.OrdinalIgnoreCase)
                            && detail.Contains("rowIds=2", StringComparison.OrdinalIgnoreCase)));
                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Speaker notes" && feature.Count == 1);
                    Assert.Contains(report.EditableFeatures, feature => feature.Name == "Slide transitions" && feature.Count == 1);
                    Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Charts" && feature.Count == 1);
                    Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Images" && feature.Count == 1);
                    Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "External relationships" && feature.Count == 1);
                    Assert.DoesNotContain(report.PreservedFeatures, feature => feature.Name == "Embedded packages");
                    Assert.Empty(report.UnsupportedFeatures);
                    Assert.Same(report, report.EnsureNoUnsupportedFeatures());
                    Assert.Contains("| Content | Tables |", report.ToMarkdown());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointTableCells_IncludeLanguageAwareRunDefaults() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointTable table = presentation.AddSlide().AddTable(1, 1);
                    table.GetCell(0, 0).Text = "Header";
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, false)) {
                    A.TableCell cell = document.PresentationPart!.SlideParts.First().Slide.Descendants<A.TableCell>().First();
                    A.Paragraph paragraph = cell.TextBody!.Elements<A.Paragraph>().First();
                    A.RunProperties runProperties = paragraph.Elements<A.Run>().First().RunProperties!;
                    A.EndParagraphRunProperties endProperties = paragraph.GetFirstChild<A.EndParagraphRunProperties>()!;

                    Assert.Equal("en-US", runProperties.Language?.Value);
                    Assert.Equal("en-US", endProperties.Language?.Value);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointTableCellText_ReplacesAllExistingParagraphs() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointTable table = presentation.AddSlide().AddTable(1, 1);
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    A.TableCell cell = document.PresentationPart!.SlideParts.First().Slide.Descendants<A.TableCell>().First();
                    cell.TextBody!.Append(
                        new A.Paragraph(new A.Run(new A.Text("Stale one"))),
                        new A.Paragraph(new A.Run(new A.Text("Stale two"))));
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointTableCell cell = presentation.Slides.Single().Tables.Single().GetCell(0, 0);
                    cell.Text = "Fresh";
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointTableCell cell = presentation.Slides.Single().Tables.Single().GetCell(0, 0);
                    Assert.Equal("Fresh", cell.Text);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DoesNotTreatChartWorkbookAsEmbeddedPackage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddChart();
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Charts" && feature.Count == 1
                        && feature.Details.Any(detail => detail.Contains("Microsoft_Excel_Worksheet", StringComparison.OrdinalIgnoreCase)));
                    Assert.DoesNotContain(report.PreservedFeatures, feature => feature.Name == "Embedded packages");
                    Assert.False(report.HasAdvancedFeatures);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DoesNotTreatOfficeImoMediaTimingAsAdvancedAnimation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    using var media = new MemoryStream(new byte[] { 1, 2, 3, 4, 5 });
                    presentation.AddSlide().AddAudio(media, "audio/mpeg", ".mp3");
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Audio and video" && feature.Count == 1);
                    Assert.DoesNotContain(report.PreservedFeatures, feature => feature.Name == "Animations and timing");
                    Assert.False(report.HasAdvancedFeatures);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsAdvancedPackageSignals() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Advanced package signals");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    PresentationPart presentationPart = document.PresentationPart!;
                    CustomXmlPart customXmlPart = presentationPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                    using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("<root><value>42</value></root>"))) {
                        customXmlPart.FeedData(stream);
                    }

                    AddExtendedPart(presentationPart,
                        "http://schemas.microsoft.com/office/2006/relationships/vbaProject",
                        "application/vnd.ms-office.vbaProject",
                        new byte[] { 1, 2, 3, 4 });
                    AddExtendedPart(presentationPart,
                        "http://schemas.microsoft.com/office/2011/relationships/webextension",
                        "application/vnd.ms-office.webextension+xml",
                        "<we:webextension xmlns:we=\"http://schemas.microsoft.com/office/webextensions/webextension/2010/11\" />");
                    AddExtendedPart(presentationPart,
                        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature",
                        "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml",
                        "<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\" />");
                    document.AddDigitalSignatureOriginPart();
                    XmlSignaturePart signaturePart = document.DigitalSignatureOriginPart!.AddNewPart<XmlSignaturePart>();
                    using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("<Signature xmlns=\"http://www.w3.org/2000/09/xmldsig#\" />"))) {
                        signaturePart.FeedData(stream);
                    }
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();

                    PowerPointFeatureFinding customXml = Assert.Single(report.FindFeatures("Custom XML parts"));
                    PowerPointFeatureFinding macros = Assert.Single(report.FindFeatures("VBA macros"));
                    PowerPointFeatureFinding webExtensions = Assert.Single(report.FindFeatures("Web extensions and task panes"));
                    PowerPointFeatureFinding signatures = Assert.Single(report.FindFeatures("Digital signatures"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, customXml.SupportLevel);
                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, macros.SupportLevel);
                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, webExtensions.SupportLevel);
                    Assert.Equal(PowerPointFeatureSupportLevel.Unsupported, signatures.SupportLevel);
                    Assert.Contains(macros.Details, detail => detail.Contains("vbaProject", StringComparison.OrdinalIgnoreCase));
                    Assert.Contains(webExtensions.Details, detail => detail.Contains("webextension", StringComparison.OrdinalIgnoreCase));
                    Assert.Contains(signatures.Details, detail => detail.Contains("signature", StringComparison.OrdinalIgnoreCase));

                    InvalidOperationException advancedException = Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                    Assert.Contains("Custom XML parts", advancedException.Message);
                    Assert.Contains("Digital signatures", advancedException.Message);

                    InvalidOperationException unsupportedException = Assert.Throws<InvalidOperationException>(() => report.EnsureNoUnsupportedFeatures());
                    Assert.Contains("Digital signatures", unsupportedException.Message);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsExternalRelationshipsWithoutHyperlinks() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Linked asset");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    document.PresentationPart!.AddExternalRelationship(
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                        new Uri("https://example.com/logo.png"));
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding external = Assert.Single(report.FindFeatures("External relationships"));

                    Assert.Equal(PowerPointFeatureSupportLevel.PartiallyEditable, external.SupportLevel);
                    Assert.Equal(1, external.Count);
                    Assert.Contains(external.Details, detail => detail.Contains("relationships/image", StringComparison.OrdinalIgnoreCase)
                        && detail.Contains("https://example.com/logo.png", StringComparison.OrdinalIgnoreCase));
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void PowerPointFeatureReport_DetectsUnsupportedTransitionMarkup() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide().AddTextBox("Unsupported transition");
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    SlidePart slidePart = document.PresentationPart!.SlideParts.Single();
                    slidePart.Slide.Transition = new Transition(
                        new OpenXmlUnknownElement("p14", "doors", "http://schemas.microsoft.com/office/powerpoint/2010/main"));
                    slidePart.Slide.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Open(filePath)) {
                    PowerPointFeatureReport report = presentation.InspectFeatures();
                    PowerPointFeatureFinding unsupported = Assert.Single(report.FindFeatures("Unsupported transition markup"));

                    Assert.Equal(PowerPointFeatureSupportLevel.Preserved, unsupported.SupportLevel);
                    Assert.Equal(1, unsupported.Count);
                    Assert.Contains(unsupported.Details, detail => detail.Contains("doors", StringComparison.OrdinalIgnoreCase));
                    Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static void AddExtendedPart(OpenXmlPartContainer container, string relationshipType, string contentType, string xml) {
            AddExtendedPart(container, relationshipType, contentType, Encoding.UTF8.GetBytes(xml));
        }

        private static void AddExtendedPart(OpenXmlPartContainer container, string relationshipType, string contentType, byte[] bytes) {
            ExtendedPart part = container.AddExtendedPart(relationshipType, contentType, "xml");
            using var stream = new MemoryStream(bytes);
            part.FeedData(stream);
        }
    }
}
