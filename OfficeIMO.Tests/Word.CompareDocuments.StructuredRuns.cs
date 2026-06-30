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

        [Fact]
        public void CompareStructureReportsRunFormattingWhenUnchangedTextIsResegmented() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_resegmented_run_format.docx");
            CreateDocumentWithRuns(sourcePath, ("A", true), ("B", false));

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_resegmented_run_format.docx");
            CreateDocumentWithRuns(targetPath, ("AB", false));

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.DoesNotContain(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Paragraph &&
                finding.Message == "Paragraph text changed.");

            WordComparisonFinding run = Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Run &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.Message == "Run formatting changed.");
            Assert.Equal("paragraph[0]/run[0]", run.Location);
            Assert.Equal("AB", run.SourceText);
            Assert.Equal("AB", run.TargetText);
        }

        private static void CreateDocumentWithSplitRuns(string path, string firstText, string secondText, bool boldSecondRun) {
            CreateDocumentWithRuns(path, (firstText, false), (secondText, boldSecondRun));
        }

        private static void CreateDocumentWithRuns(string path, params (string Text, bool Bold)[] runs) {
            using WordDocument doc = WordDocument.Create(path);
            doc.AddParagraph("Placeholder");
            doc.Save(false);

            using WordprocessingDocument document = WordprocessingDocument.Open(path, true);
            Body body = document.MainDocumentPart!.Document.Body!;
            var paragraph = new Paragraph();
            foreach ((string text, bool bold) in runs) {
                var run = new Run(new Text(text));
                if (bold) {
                    run.RunProperties = new RunProperties(new Bold());
                }

                paragraph.Append(run);
            }

            body.RemoveAllChildren<Paragraph>();
            body.PrependChild(paragraph);
            document.MainDocumentPart.Document.Save();
        }
    }
}
