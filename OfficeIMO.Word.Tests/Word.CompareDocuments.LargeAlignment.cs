using System.IO;
using System.Linq;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareStructureAlignsRepeatedParagraphsBeyondAMillionAlignmentCells() {
            // 1,101 x 1,101 paragraphs is more than a million alignment cells, where a greedy first-match
            // pass pairs the moved paragraph with the last line and reports every repeated line as changed.
            const int repeated = 1_100;
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_source_large_repeated.docx");
            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Moved paragraph");
                for (int index = 0; index < repeated; index++) doc.AddParagraph("Repeated line");
                doc.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_target_large_repeated.docx");
            using (WordDocument doc = WordDocument.Create(targetPath)) {
                for (int index = 0; index < repeated; index++) doc.AddParagraph("Repeated line");
                doc.AddParagraph("Moved paragraph");
                doc.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            var paragraphChanges = result.Findings
                .Where(finding => finding.Scope == WordComparisonScope.Paragraph && finding.ChangeKind != WordComparisonChangeKind.Modified)
                .ToList();
            Assert.Equal(2, paragraphChanges.Count);
            Assert.Contains(paragraphChanges, finding => finding.ChangeKind == WordComparisonChangeKind.Deleted && finding.SourceText == "Moved paragraph");
            Assert.Contains(paragraphChanges, finding => finding.ChangeKind == WordComparisonChangeKind.Inserted && finding.TargetText == "Moved paragraph");
        }
    }
}
