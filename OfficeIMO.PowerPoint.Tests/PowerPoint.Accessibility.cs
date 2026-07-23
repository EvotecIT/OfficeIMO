using System;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests {
    public class PowerPointAccessibilityTests {
        [Fact]
        public void ShapeAccessibilityMetadataAndReadingOrderRoundTrip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    PowerPointSlide slide = presentation.AddSlide();
                    PowerPointAutoShape accent = slide.AddRectanglePoints(10, 10, 40, 20, "Decorative accent");
                    accent.Title = "Decorative accent";
                    accent.Decorative = true;
                    PowerPointTextBox text = slide.AddTextBoxPoints("Accessible documentation", 20, 40, 220, 30);
                    text.Title = "Slide title";
                    text.Description = "Documentation link";
                    text.SetLanguage("en-GB");
                    PowerPointTextRun run = text.Paragraphs.Single().Runs.Single();
                    run.SetHyperlink("https://example.test/docs", "Open the documentation");

                    Assert.Equal(1, text.ReadingOrder);
                    text.MoveToReadingOrder(0);
                    Assert.Equal(0, text.ReadingOrder);
                    Assert.Equal(1, accent.ReadingOrder);
                    Assert.Equal("en-GB", text.Language);
                    Assert.Equal("Open the documentation", run.HyperlinkTooltip);
                    Assert.True(run.HasMeaningfulHyperlinkLabel);
                    Assert.Empty(presentation.ValidateDocument());
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointTextBox text = presentation.Slides[0].TextBoxes.Single();
                    PowerPointAutoShape accent = presentation.Slides[0].Shapes.OfType<PowerPointAutoShape>().Single();
                    Assert.Equal("Slide title", text.Title);
                    Assert.Equal("Documentation link", text.Description);
                    Assert.Equal("en-GB", text.Language);
                    Assert.True(accent.Decorative);
                    Assert.Equal(0, text.ReadingOrder);
                    Assert.Equal("Open the documentation",
                        text.Paragraphs.Single().Runs.Single().HyperlinkTooltip);
                }
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void StrictAccessibilityProfileReturnsStructuredPolicyFindings() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            PowerPointSlide slide = presentation.AddSlide();
            slide.BackgroundColor = "FFFFFF";
            PowerPointTextBox link = slide.AddTextBoxPoints("click here", 20, 20, 180, 28);
            link.Color = "D0D0D0";
            link.FontSize = 12;
            link.FillColor = "FFFFFF";
            link.Paragraphs.Single().Runs.Single().SetHyperlink("https://example.test");
            PowerPointTable table = slide.AddTablePoints(2, 2, 20, 70, 180, 60);
            table.GetCell(0, 0).Text = "Metric";
            table.GetCell(0, 1).Text = "Value";
            table.GetCell(1, 0).Text = "Users";
            table.GetCell(1, 1).Text = "42";
            table.HeaderRow = false;
            var chartData = new PowerPointChartData(new[] { "Q1", "Q2" }, new[] {
                new PowerPointChartSeries("Actual", new[] { 10D, 15D }),
                new PowerPointChartSeries("Target", new[] { 12D, 18D })
            });
            slide.AddChartPoints(chartData, 220, 70, 240, 150);

            PowerPointAccessibilityReport report = presentation.InspectAccessibility(
                PowerPointAccessibilityOptions.ForProfile(PowerPointAccessibilityPolicyProfile.Strict));

            Assert.False(report.IsSuccessful);
            Assert.Contains(report.Findings, finding => finding.Code == "Accessibility.MissingDocumentTitle");
            Assert.Contains(report.Findings, finding => finding.Code == "Accessibility.MissingShapeTitle");
            Assert.Contains(report.Findings, finding => finding.Code == "Accessibility.MissingAlternativeText");
            Assert.Contains(report.Findings, finding => finding.Code == "Accessibility.MissingLanguage");
            Assert.Contains(report.Findings, finding => finding.Code == "Accessibility.MissingTableHeader");
            Assert.Contains(report.Findings, finding => finding.Code == "Accessibility.LowContrast" &&
                finding.MeasuredValue < finding.RequiredValue);
            Assert.Contains(report.Findings, finding => finding.Code == "Accessibility.UnclearLinkLabel");
            Assert.Contains(report.Findings, finding => finding.Code == "Accessibility.ChartColorOnlyMeaning");
            Assert.Throws<PowerPointAccessibilityException>(() => report.EnsureCompliant());
            Assert.Contains("\"profile\":\"Strict\"", report.ToJson(), StringComparison.Ordinal);
        }

        [Fact]
        public void AccessibilityGroupsContiguousRunsWithTheSameHyperlinkBeforeJudgingTheLabel() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            PowerPointTextBox link = presentation.AddSlide().AddTextBoxPoints("Open", 20, 20, 180, 28);
            PowerPointParagraph paragraph = link.Paragraphs.Single();
            paragraph.Runs.Single().SetHyperlink("https://openai.com");
            paragraph.AddRun("AI", run => {
                run.Bold = true;
                run.SetHyperlink("https://openai.com");
            });

            PowerPointAccessibilityReport report = presentation.InspectAccessibility();

            Assert.DoesNotContain(report.Findings,
                finding => finding.Code == "Accessibility.UnclearLinkLabel");
        }

        [Fact]
        public void DesignerSlidesPassDefaultAccessibilityProfileWithoutCallerCleanup() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            presentation.AddDesignerSectionSlide("Delivery and evidence", "Accessible by default");
            presentation.AddDesignerProcessSlide("A controlled workflow", "Every step remains editable", new[] {
                new PowerPointProcessStep("Inspect", "Read the source"),
                new PowerPointProcessStep("Generate", "Create native content"),
                new PowerPointProcessStep("Validate", "Enforce the contract")
            });
            var tableData = new PowerPointTableData(new[] { "Metric", "Value" }, new[] {
                new[] { "Quality", "Pass" },
                new[] { "Coverage", "Complete" }
            });
            presentation.AddDesignerAppendixTableSlide("Evidence appendix", null, tableData);

            PowerPointAccessibilityReport report = presentation.InspectAccessibility();

            Assert.True(report.IsSuccessful, string.Join(Environment.NewLine,
                report.Findings.Where(finding => finding.Severity == PowerPointAccessibilitySeverity.Error)
                    .Select(finding => finding.Code + " [slide " + finding.SlideIndex + ", " +
                        finding.ShapeName + ", " + report.Slides[finding.SlideIndex ?? 0].Shapes
                            .FirstOrDefault(shape => shape.ShapeId == finding.ShapeId)?.Title + "]: " +
                        finding.Message)));
            Assert.All(report.Slides, slide => Assert.False(string.IsNullOrWhiteSpace(slide.Title)));
            Assert.All(report.Slides.SelectMany(slide => slide.Shapes)
                .Where(shape => shape.ContentType == PowerPointShapeContentType.TextBox),
                shape => Assert.False(string.IsNullOrWhiteSpace(shape.Language)));
        }

        [Fact]
        public void AccessibilitySkipsHiddenSlidesUnlessExplicitlyIncluded() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            PowerPointSlide hidden = presentation.AddSlide();
            hidden.Hidden = true;
            hidden.AddRectanglePoints(20, 20, 120, 80, "Undescribed visual");

            PowerPointAccessibilityReport defaultReport = presentation.InspectAccessibility();
            PowerPointAccessibilityReport includedReport = presentation.InspectAccessibility(
                new PowerPointAccessibilityOptions { IncludeHiddenSlides = true });

            Assert.Empty(defaultReport.Slides);
            Assert.DoesNotContain(defaultReport.Findings, finding => finding.SlideIndex == 0);
            Assert.Single(includedReport.Slides);
            Assert.Contains(includedReport.Findings,
                finding => finding.SlideIndex == 0 && finding.Code == "Accessibility.MissingSlideTitle");
        }

        [Fact]
        public void AccessibilityDoesNotTreatInheritedPlaceholderPromptAsAuthoredSlideTitle() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTextBox inheritedTitle = presentation.EnsureLayoutPlaceholderTextBox(0, slide.LayoutIndex,
                PlaceholderValues.Title, bounds: PowerPointLayoutBox.FromCentimeters(1D, 1D, 20D, 2D));
            inheritedTitle.Text = "Click to add title";

            PowerPointAccessibilityReport report = presentation.InspectAccessibility();

            Assert.Contains(report.Findings, finding =>
                finding.SlideIndex == 0 && finding.Code == "Accessibility.MissingSlideTitle");
            Assert.Null(Assert.Single(report.Slides).Title);
        }

        [Fact]
        public void AccessibilityInspectsTableNestedInsideGroup() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointTable table = slide.AddTablePoints(2, 2, 20, 50, 180, 70);
            table.HeaderRow = false;
            PowerPointAutoShape anchor = slide.AddRectanglePoints(210, 50, 20, 20, "Group anchor");
            slide.GroupShapes(new PowerPointShape[] { table, anchor }, "Accessibility group");

            PowerPointAccessibilityReport report = presentation.InspectAccessibility(
                new PowerPointAccessibilityOptions {
                    RequireSlideTitles = false,
                    RequireAlternativeText = false,
                    RequireLanguage = false,
                    CheckContrast = false,
                    CheckMeaningfulLinks = false,
                    CheckColorOnlyMeaning = false
                });

            Assert.Contains(report.Findings, finding =>
                finding.Code == "Accessibility.MissingTableHeader" && finding.ShapeId == table.Id);
            Assert.Contains(Assert.Single(report.Slides).Shapes, shape => shape.ShapeId == table.Id);
        }

        [Fact]
        public void AccessibilityRejectsShapeTreesBeyondTheConfiguredLimit() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddRectanglePoints(10, 10, 20, 20, "One");
            slide.AddRectanglePoints(40, 10, 20, 20, "Two");

            Assert.Throws<InvalidOperationException>(() => presentation.InspectAccessibility(
                new PowerPointAccessibilityOptions { MaximumShapeCount = 1 }));
            Assert.Throws<InvalidOperationException>(() => presentation.InspectPreflight(
                new PowerPointDeckPreflightOptions { MaximumShapeCount = 1 }));
        }

        [Fact]
        public void AccessibilityBoundsGroupChildrenWhileMaterializingTheTree() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointAutoShape first = slide.AddRectanglePoints(10, 10, 20, 20, "One");
            PowerPointAutoShape second = slide.AddRectanglePoints(40, 10, 20, 20, "Two");
            PowerPointGroupShape group = slide.GroupShapes(new PowerPointShape[] { first, second }, "Bounded group");

            Assert.Throws<InvalidOperationException>(() => slide.GetGroupChildren(group, maximumChildren: 1));
            Assert.Throws<InvalidOperationException>(() => presentation.InspectAccessibility(
                new PowerPointAccessibilityOptions { MaximumShapeCount = 2 }));
            Assert.Throws<InvalidOperationException>(() => presentation.InspectPreflight(
                new PowerPointDeckPreflightOptions { MaximumShapeCount = 2 }));
        }

        [Fact]
        public void PackageFingerprintBoundsRelationshipFanOutWhileCollectingParts() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            presentation.AddSlide();

            Assert.ThrowsAny<IOException>(() => PowerPointPackageFingerprint.CollectParts(
                presentation.OpenXmlDocument,
                maximumPartCount: 100,
                maximumPartDepth: 100,
                maximumRelationshipCount: 1));
        }

        [Fact]
        public void PackageFingerprintPreservesUnsavedSlideRoots() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddRectanglePoints(10, 10, 20, 20, "Unsaved shape");

            _ = PowerPointPackageFingerprint.Create(presentation.OpenXmlDocument);

            Assert.Single(slide.Shapes);
            Assert.Empty(slide.ClassicAnimations);
        }

        [Fact]
        public void PackageFingerprintRejectsOversizedXmlStreamsBeforeParsing() {
            using var stream = new MemoryStream(new byte[17]);

            Assert.ThrowsAny<IOException>(() =>
                PowerPointPackageFingerprint.EnsureStreamWithinLimit(stream, maximumBytes: 16));
        }

        [Fact]
        public void PackageFingerprintChecksApplicationPropertiesSizeBeforeLoadingSignatureMetadata() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide();
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    ExtendedFilePropertiesPart propertiesPart = document.ExtendedFilePropertiesPart
                        ?? document.AddExtendedFilePropertiesPart();
                    propertiesPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                    propertiesPart.Properties.DigitalSignature = new DocumentFormat.OpenXml.ExtendedProperties.DigitalSignature();
                    propertiesPart.Properties.Save();
                }

                using PresentationDocument reopened = PresentationDocument.Open(filePath, false);
                ExtendedFilePropertiesPart reopenedProperties = reopened.ExtendedFilePropertiesPart!;
                Assert.False(reopenedProperties.IsRootElementLoaded);
                long serializedLength;
                using (Stream source = reopenedProperties.GetStream(FileMode.Open, FileAccess.Read)) {
                    serializedLength = source.Length;
                }

                Assert.ThrowsAny<IOException>(() =>
                    PowerPointPackageFingerprint.HasApplicationSignatureFlag(reopened, serializedLength - 1));
                Assert.False(reopenedProperties.IsRootElementLoaded);
                Assert.True(PowerPointPackageFingerprint.HasApplicationSignatureFlag(reopened, serializedLength));
                Assert.True(reopenedProperties.IsRootElementLoaded);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void SignedPackageXmlIsBoundedBeforePresentationRootsAreLoaded() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            try {
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddSlide();
                    presentation.Save();
                }

                using (PresentationDocument document = PresentationDocument.Open(filePath, true)) {
                    ExtendedFilePropertiesPart propertiesPart = document.ExtendedFilePropertiesPart
                        ?? document.AddExtendedFilePropertiesPart();
                    propertiesPart.Properties ??= new DocumentFormat.OpenXml.ExtendedProperties.Properties();
                    propertiesPart.Properties.DigitalSignature = new DocumentFormat.OpenXml.ExtendedProperties.DigitalSignature();
                    propertiesPart.Properties.Save();

                    PresentationPart presentationPart = document.PresentationPart!;
                    string xml;
                    using (var reader = new StreamReader(
                               presentationPart.GetStream(FileMode.Open, FileAccess.Read),
                               Encoding.UTF8, detectEncodingFromByteOrderMarks: true)) {
                        xml = reader.ReadToEnd();
                    }
                    xml = xml.Replace(
                        "</p:presentation>",
                        "<!--" + new string('x', 4096) + "--></p:presentation>",
                        StringComparison.Ordinal);
                    using (var writer = new StreamWriter(
                               presentationPart.GetStream(FileMode.Create, FileAccess.Write),
                               new UTF8Encoding(encoderShouldEmitUTF8Identifier: false))) {
                        writer.Write(xml);
                    }
                }

                using PresentationDocument reopened = PresentationDocument.Open(filePath, false);
                PresentationPart reopenedPresentation = reopened.PresentationPart!;
                Assert.False(reopenedPresentation.IsRootElementLoaded);
                long applicationPropertiesLength;
                using (Stream source = reopened.ExtendedFilePropertiesPart!
                           .GetStream(FileMode.Open, FileAccess.Read)) {
                    applicationPropertiesLength = source.Length;
                }
                long presentationLength;
                using (Stream source = reopenedPresentation.GetStream(FileMode.Open, FileAccess.Read)) {
                    presentationLength = source.Length;
                }
                Assert.True(presentationLength > applicationPropertiesLength);

                Assert.ThrowsAny<IOException>(() =>
                    PowerPointPackageFingerprint.HasSignatureAndEnsurePackageXmlWithinLimit(
                        reopened, presentationLength - 1));
                Assert.False(reopenedPresentation.IsRootElementLoaded);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void AccessibilityReportIsStableForGeneratedAndReloadedDecks() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");
            string reportPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".json");
            try {
                string[] generatedCodes;
                using (PowerPointPresentation presentation = PowerPointPresentation.Create(filePath)) {
                    presentation.AddDesignerSectionSlide("Portable evidence", "Generated and imported inspection agree");
                    PowerPointAccessibilityReport generated = presentation.InspectAccessibility();
                    generatedCodes = generated.Findings.Select(finding => finding.Code).ToArray();
                    generated.SaveJson(reportPath);
                    presentation.Save();
                }

                using (PowerPointPresentation presentation = PowerPointPresentation.Load(filePath, new PowerPointLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                    PowerPointAccessibilityReport imported = presentation.InspectAccessibility();
                    Assert.Equal(generatedCodes, imported.Findings.Select(finding => finding.Code));
                    Assert.True(imported.IsSuccessful);
                }
                Assert.Contains("\"schemaVersion\":1", File.ReadAllText(reportPath), StringComparison.Ordinal);
                using PresentationDocument document = PresentationDocument.Open(filePath, false);
                Assert.Empty(new OpenXmlValidator().Validate(document));
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
                if (File.Exists(reportPath)) File.Delete(reportPath);
            }
        }

        [Fact]
        public void AccessibilityContrastUsesInheritedLayoutShapeBackground() {
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream, new PowerPointCreateOptions());
            PowerPointSlide slide = presentation.AddSlide();
            SlideLayoutPart layoutPart = slide.SlidePart.SlideLayoutPart!;
            DocumentFormat.OpenXml.Presentation.ShapeTree tree =
                layoutPart.SlideLayout.CommonSlideData!.ShapeTree!;
            tree.AppendChild(new DocumentFormat.OpenXml.Presentation.Shape(
                new DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
                    new DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties {
                        Id = 900U, Name = "Inherited dark panel"
                    },
                    new DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(),
                    new DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()),
                new DocumentFormat.OpenXml.Presentation.ShapeProperties(
                    new DocumentFormat.OpenXml.Drawing.Transform2D(
                        new DocumentFormat.OpenXml.Drawing.Offset {
                            X = PowerPointUnits.FromPoints(10), Y = PowerPointUnits.FromPoints(10)
                        },
                        new DocumentFormat.OpenXml.Drawing.Extents {
                            Cx = PowerPointUnits.FromPoints(220), Cy = PowerPointUnits.FromPoints(70)
                        }),
                    new DocumentFormat.OpenXml.Drawing.PresetGeometry(
                        new DocumentFormat.OpenXml.Drawing.AdjustValueList()) {
                            Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
                        },
                    new DocumentFormat.OpenXml.Drawing.SolidFill(
                        new DocumentFormat.OpenXml.Drawing.RgbColorModelHex { Val = "111827" }))));
            layoutPart.SlideLayout.Save();

            PowerPointTextBox text = slide.AddTextBoxPoints("Readable inherited contrast", 20, 20, 180, 30);
            text.Color = "FFFFFF";
            text.FontSize = 12;

            PowerPointAccessibilityReport report = presentation.InspectAccessibility(
                new PowerPointAccessibilityOptions {
                    RequireSlideTitles = false,
                    RequireLanguage = false,
                    RequireAlternativeText = false,
                    RequireTableHeaders = false,
                    CheckMeaningfulLinks = false,
                    CheckColorOnlyMeaning = false
                });

            Assert.DoesNotContain(report.Findings, finding =>
                finding.Code == "Accessibility.LowContrast" && finding.ShapeId == text.Id);
        }
    }
}
