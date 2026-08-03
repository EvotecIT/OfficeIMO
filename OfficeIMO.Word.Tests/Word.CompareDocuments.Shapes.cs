using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareStructure_ReportsDrawingMlShapeContentAndGroupInsertion() {
            string sourcePath = Path.Combine(_directoryWithFiles, "Compare.Shape.Source.docx");
            string targetPath = Path.Combine(_directoryWithFiles, "Compare.Shape.Target.docx");
            CreateShapeComparisonPair(sourcePath, targetPath);

            WordComparisonResult result = WordDocumentComparer.CompareStructure(
                sourcePath,
                targetPath,
                new WordComparisonOptions { CompareGeneratedIds = false });

            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Shape &&
                finding.ChangeKind == WordComparisonChangeKind.Modified &&
                finding.Message == "DrawingML shape content changed.");
            Assert.Contains(result.Findings, finding =>
                finding.Scope == WordComparisonScope.Shape &&
                finding.ChangeKind == WordComparisonChangeKind.Inserted &&
                finding.TargetText!.Contains("shape-group", System.StringComparison.Ordinal));
            Assert.All(result.Findings.Where(finding => finding.Scope == WordComparisonScope.Shape), finding =>
                Assert.StartsWith("body/shape[", finding.DetailedLocation, System.StringComparison.Ordinal));

            WordComparisonResult disabled = WordDocumentComparer.CompareStructure(
                sourcePath,
                targetPath,
                new WordComparisonOptions { CompareShapes = false, CompareGeneratedIds = false });
            Assert.DoesNotContain(disabled.Findings, finding => finding.Scope == WordComparisonScope.Shape);
        }

        [Fact]
        public void CreateRedlineDocument_TracksShapeEvidenceAndDisclosesInPlaceFallback() {
            string sourcePath = Path.Combine(_directoryWithFiles, "Redline.Shape.Source.docx");
            string targetPath = Path.Combine(_directoryWithFiles, "Redline.Shape.Target.docx");
            string reportPath = Path.Combine(_directoryWithFiles, "Redline.Shape.Report.docx");
            string inPlacePath = Path.Combine(_directoryWithFiles, "Redline.Shape.InPlace.docx");
            CreateShapeComparisonPair(sourcePath, targetPath);

            WordComparisonResult report = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                reportPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.ReportArtifact,
                    ComparisonOptions = new WordComparisonOptions { CompareGeneratedIds = false }
                });
            Assert.Contains(report.Findings, finding => finding.Scope == WordComparisonScope.Shape);
            using (WordDocument artifact = WordDocument.Load(reportPath)) {
                Assert.Contains(artifact.Paragraphs, paragraph =>
                    paragraph.Text.Contains("shape-group", System.StringComparison.Ordinal));
                Assert.Empty(new OpenXmlValidator().Validate(artifact._wordprocessingDocument));
            }

            WordComparisonResult inPlace = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                inPlacePath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    ComparisonOptions = new WordComparisonOptions { CompareGeneratedIds = false }
                });
            Assert.Contains(inPlace.Limitations, limitation =>
                limitation.Code == "Redline.Shape.InPlaceTextFallback");
            using WordDocument redline = WordDocument.Load(inPlacePath);
            Assert.Contains(redline.Paragraphs, paragraph =>
                paragraph.Text.Contains("Shape Modified", System.StringComparison.Ordinal));
            Assert.Contains(redline._wordprocessingDocument.MainDocumentPart!.Document.Body!
                .Descendants<InsertedRun>(), revision => revision.InnerText.Contains("shape", System.StringComparison.Ordinal));
            Assert.Empty(new OpenXmlValidator().Validate(redline._wordprocessingDocument));
        }

        [Fact]
        public void CompareStructure_IgnoresGeneratedShapeGroupIdsOnlyWhenRequested() {
            string sourcePath = Path.Combine(_directoryWithFiles, "Compare.ShapeGroup.GeneratedIds.Source.docx");
            string targetPath = Path.Combine(_directoryWithFiles, "Compare.ShapeGroup.GeneratedIds.Target.docx");
            CreateIdenticalShapeGroupDocument(sourcePath);
            CreateIdenticalShapeGroupDocument(targetPath);

            WordComparisonResult suppressed = WordDocumentComparer.CompareStructure(
                sourcePath,
                targetPath,
                new WordComparisonOptions { CompareGeneratedIds = false });
            Assert.DoesNotContain(suppressed.Findings, finding => finding.Scope == WordComparisonScope.Shape);

            WordComparisonResult included = WordDocumentComparer.CompareStructure(
                sourcePath,
                targetPath,
                new WordComparisonOptions { CompareGeneratedIds = true });
            Assert.Contains(included.Findings, finding =>
                finding.Scope == WordComparisonScope.Shape &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);
        }

        private static void CreateShapeComparisonPair(string sourcePath, string targetPath) {
            using (WordDocument source = WordDocument.Create(sourcePath)) {
                source.AddParagraph().AddShapeDrawing(ShapeType.Rectangle, 100, 50);
                source.Save();
            }
            using (WordDocument target = WordDocument.Create(targetPath)) {
                target.AddParagraph().AddShapeDrawing(ShapeType.Ellipse, 100, 50);
                target.AddParagraph().AddShapeGroup(new[] {
                    new WordShapeGroupItem(ShapeType.Chevron, 0, 0, 60, 30),
                    new WordShapeGroupItem(ShapeType.Chevron, 48, 0, 60, 30)
                });
                target.Save();
            }
        }

        private static void CreateIdenticalShapeGroupDocument(string path) {
            using WordDocument document = WordDocument.Create(path);
            document.AddParagraph().AddShapeGroup(new[] {
                new WordShapeGroupItem(ShapeType.Chevron, 0, 0, 60, 30),
                new WordShapeGroupItem(ShapeType.Chevron, 48, 0, 60, 30)
            });
            document.Save();
        }
    }
}
