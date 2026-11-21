using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_InspectAccessibility_FindsEachRuleAndFixAccessibilityPersistsSafeFixes() {
            string filePath = Path.Combine(_directoryWithFiles, "Accessibility.Rules.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddHeading("Overview", 1);
                document.AddHeading("Details", 3);
                WordImage inline = document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"));
                inline.AltText = null;
                inline.Title = null;
                WordImage floating = document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "Kulek.jpg"));
                floating.AltText = "A cat sleeping on a sofa.";
                floating.WrapText = WordImageTextWrapping.Square;
                WordTable data = document.AddTable(3, 2);
                data.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
                document.AddTable(1, 3);
                document.AddHyperLink("click here", new Uri("https://evotec.xyz"));
                document.AddParagraph("Pale text").SetColor(OfficeColor.Parse("#DDDDDD"));
                document.AddParagraph();
                document.AddParagraph();
                document.AddParagraph("After the gap");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                OfficeAccessibilityReport report = document.InspectAccessibility();
                var codes = report.Findings.Select(finding => finding.Code).ToList();
                Assert.Contains(OfficeDiagnosticCodes.AccessibilityMissingDocumentTitle, codes);
                Assert.Contains(OfficeDiagnosticCodes.AccessibilityHeadingLevelSkipped, codes);
                Assert.Contains(OfficeDiagnosticCodes.AccessibilityMissingAlternativeText, codes);
                Assert.Contains("accessibility-floating-object", codes);
                Assert.Contains(OfficeDiagnosticCodes.AccessibilityMissingTableHeader, codes);
                Assert.Contains(OfficeDiagnosticCodes.AccessibilityUnclearLinkLabel, codes);
                Assert.Contains(OfficeDiagnosticCodes.AccessibilityLowContrast, codes);
                Assert.Contains("accessibility-blank-paragraph-spacing", codes);
                OfficeAccessibilityFinding layout = Assert.Single(report.Findings, finding => finding.Code == "accessibility-layout-table");
                Assert.True(layout.IsHeuristic);
                Assert.Equal(OfficeAccessibilitySeverity.Warning, layout.Severity);
                OfficeAccessibilityFinding contrast = report.Findings.Single(finding => finding.Code == OfficeDiagnosticCodes.AccessibilityLowContrast);
                Assert.True(contrast.MeasuredValue < contrast.RequiredValue);
                Assert.False(report.IsSuccessful);
                Assert.Throws<OfficeAccessibilityException>(() => report.EnsureCompliant());

                OfficeAccessibilityReport fixedReport = document.FixAccessibility(new WordAccessibilityFixOptions {
                    Title = "Quarterly review",
                    Language = "en-GB",
                    MarkTableHeaderRows = true
                });
                Assert.DoesNotContain(fixedReport.Findings, finding => finding.Code == OfficeDiagnosticCodes.AccessibilityMissingDocumentTitle);
                Assert.DoesNotContain(fixedReport.Findings, finding => finding.Code == OfficeDiagnosticCodes.AccessibilityMissingTableHeader);
                Assert.Contains(fixedReport.Findings, finding => finding.Code == "accessibility-layout-table");
                document.Save();
            }

            using (WordDocument reopened = WordDocument.Load(filePath)) {
                Assert.Empty(reopened.ValidateDocument());
                OfficeAccessibilityReport report = reopened.InspectAccessibility();
                Assert.DoesNotContain(report.Findings, finding => finding.Code == OfficeDiagnosticCodes.AccessibilityMissingDocumentTitle);
                Assert.DoesNotContain(report.Findings, finding => finding.Code == OfficeDiagnosticCodes.AccessibilityMissingLanguage);
                Assert.DoesNotContain(report.Findings, finding => finding.Code == OfficeDiagnosticCodes.AccessibilityMissingTableHeader);
                Assert.Equal("Quarterly review", reopened.Properties.Title);
            }
        }

        [Fact]
        public void Test_InspectAccessibility_ReportsMissingLanguageAndHeadingsInLongDocument() {
            using WordDocument document = WordDocument.Create();
            document.Properties.Title = "Notes";
            for (int index = 0; index < 25; index++) document.AddParagraph("Paragraph " + index);
            var styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            foreach (var language in styles.Descendants<DocumentFormat.OpenXml.Wordprocessing.Languages>().ToList()) language.Remove();

            OfficeAccessibilityReport report = document.InspectAccessibility(new WordAccessibilityOptions { Profile = OfficeAccessibilityPolicyProfile.Strict });

            Assert.Contains(report.Findings, finding => finding.Code == OfficeDiagnosticCodes.AccessibilityMissingLanguage && finding.Severity == OfficeAccessibilitySeverity.Error);
            Assert.Contains(report.Findings, finding => finding.Code == "accessibility-missing-headings" && finding.Severity == OfficeAccessibilitySeverity.Error);

            document.FixAccessibility(new WordAccessibilityFixOptions { Language = "pl-PL" });
            Assert.DoesNotContain(document.InspectAccessibility().Findings, finding => finding.Code == OfficeDiagnosticCodes.AccessibilityMissingLanguage);
            Assert.Empty(document.ValidateDocument());
        }

        [Fact]
        public void Test_InspectAccessibility_UsesStyleBoldForTheLargeTextThreshold() {
            using WordDocument document = WordDocument.Create();
            WordParagraph label = document.AddParagraph("Section label").SetColor(OfficeColor.Parse("#888888")).SetFontSize(OfficeLength.Points(14));
            var styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new DocumentFormat.OpenXml.Wordprocessing.Style(
                new DocumentFormat.OpenXml.Wordprocessing.StyleName { Val = "Strong label" },
                new DocumentFormat.OpenXml.Wordprocessing.StyleRunProperties(new DocumentFormat.OpenXml.Wordprocessing.Bold())) {
                Type = DocumentFormat.OpenXml.Wordprocessing.StyleValues.Character, StyleId = "StrongLabel"
            });
            label._run!.RunProperties!.RunStyle = new DocumentFormat.OpenXml.Wordprocessing.RunStyle { Val = "StrongLabel" };

            Assert.DoesNotContain(document.InspectAccessibility().Findings, finding => finding.Code == OfficeDiagnosticCodes.AccessibilityLowContrast);
        }

        [Fact]
        public void Test_InspectAccessibility_UsesWordHighlightPaletteForContrast() {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph("Highlighted").SetColor(OfficeColor.Parse("#FFFFFF")).SetHighlight(WordHighlightColor.Green);

            OfficeAccessibilityFinding contrast = Assert.Single(document.InspectAccessibility().Findings, finding => finding.Code == OfficeDiagnosticCodes.AccessibilityLowContrast);
            Assert.True(contrast.MeasuredValue < 1.5D);
        }
    }
}
