using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareStructureReportsRunLevelTextChangesInsideModifiedParagraphs() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_run_text.docx");
            CreateDocumentWithSplitRuns(sourcePath, "Status: ", "Draft", boldSecondRun: true);

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_run_text.docx");
            CreateDocumentWithSplitRuns(targetPath, "Status: ", "Approved", boldSecondRun: true);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.SourceText == "Status: Draft" &&
                finding.TargetText == "Status: Approved");

            WordComparisonFinding run = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Run &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.Message == "Run text changed.");
            Assert.Equal("paragraph[0]/run[1]", run.Location);
            Assert.Equal("Draft", run.SourceText);
            Assert.Equal("Approved", run.TargetText);
        }

        [Fact]
        public void CompareStructureReportsRunFormattingChangesWithStableRunLocations() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_run_format.docx");
            CreateDocumentWithSplitRuns(sourcePath, "Decision: ", "Approved", boldSecondRun: false);

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_run_format.docx");
            CreateDocumentWithSplitRuns(targetPath, "Decision: ", "Approved", boldSecondRun: true);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.Message == "Paragraph text changed.");

            WordComparisonFinding run = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Run &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.Message == "Run formatting changed.");
            Assert.Equal("paragraph[0]/run[1]", run.Location);
            Assert.Equal("Approved", run.SourceText);
            Assert.Equal("Approved", run.TargetText);
        }

        private static void CreateDocumentWithSplitRuns(string path, string firstText, string secondText, bool boldSecondRun) {
            using WordDocument doc = WordDocument.Create(path);
            doc.AddParagraph("Placeholder");
            doc.Save(false);

            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Body body = document.MainDocumentPart!.Document.Body!;
            var secondRun = new Run(new Text(secondText));
            if (boldSecondRun) {
                secondRun.RunProperties = new RunProperties(new Bold());
            }

            body.RemoveAllChildren<Paragraph>();
            body.PrependChild(new Paragraph(
                new Run(new Text(firstText)),
                secondRun));
            document.MainDocumentPart.Document.Save();
        }
    }
}
