using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Provenance;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeProvenanceDocumentTextTests {
    [Fact]
    public async Task WordAssessmentUsesNativeTextLocationsAcrossStoriesAndPreservesBytes() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        try {
            using (var document = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document)) {
                var main = document.AddMainDocumentPart();
                main.Document = new W.Document(new W.Body(
                    Paragraph("body\u200B"), Paragraph("\u00A0"), Paragraph("\uFEFFnode"),
                    new W.Table(new W.TableRow(new W.TableCell(Paragraph("cell\u202E"))))));
                main.AddNewPart<HeaderPart>().Header = new W.Header(Paragraph("header\u200C"));
                main.AddNewPart<FooterPart>().Footer = new W.Footer(Paragraph("footer\u200D"));
                main.AddNewPart<FootnotesPart>().Footnotes = new W.Footnotes(new W.Footnote(Paragraph("note\u2060")) { Id = 1 });
                main.AddNewPart<EndnotesPart>().Endnotes = new W.Endnotes(new W.Endnote(Paragraph("endnote\uFE0F")) { Id = 1 });
            }
            byte[] before = File.ReadAllBytes(path);
            var request = new OfficeProvenanceWorkflowRequest { InputPath = path, Operation = OfficeProvenanceWorkflowOperation.Assess };
            var runner = new OfficeWorkflowRunner();
            var result = await runner.RunProvenanceAsync(request);
            Assert.True(result.Succeeded, result.Summary);
            Assert.Equal(OfficeProvenanceCheckStatus.Completed, result.Checks.TextIntegrity);
            var findings = result.Assessment!.TextIntegrity!.Findings;
            Assert.Equal(8, findings.Count);
            Assert.Contains(findings, finding => finding.CodePoint == 0x00A0);
            Assert.Contains(findings, finding => finding.CodePoint == 0xFEFF && finding.TextOffset == 0);
            foreach (string story in new[] { "Document/", "Header[1]/", "Footer[1]/", "Footnotes/", "Endnotes/" })
                Assert.Contains(findings, finding => finding.Location.StartsWith(story, StringComparison.Ordinal));
            Assert.Contains(findings, finding => finding.CodePoint == 0x202E && finding.TextOffset == 4);
            Assert.Equal(before, File.ReadAllBytes(path));
            Assert.True(OfficeProvenanceAudit.HasFindings(result));
            request.Assessment.InspectTextIntegrity = false;
            var disabled = await runner.RunProvenanceAsync(request);
            Assert.True(disabled.Succeeded, disabled.Summary);
            Assert.Equal(OfficeProvenanceCheckStatus.Disabled, disabled.Checks.TextIntegrity);
            request.Assessment.InspectTextIntegrity = true;
            request.Assessment.TextIntegrity.MaxCharacters = 2;
            var limited = await runner.RunProvenanceAsync(request);
            Assert.False(limited.Succeeded);
            Assert.Equal(OfficeProvenanceCheckStatus.Failed, limited.Checks.TextIntegrity);
        } finally { File.Delete(path); }
    }
    private static W.Paragraph Paragraph(string text) => new(new W.Run(new W.Text(text)));
}
