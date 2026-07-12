using System;
using System.IO;
using System.Text.Json;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void CompareStructureExportsJsonMarkdownAndTextSummaryReports() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_report_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Alpha | one");
                document.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_report_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Beta | two");
                document.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.Equal(WordComparisonReportWriter.ToJson(result), result.ToJson());
            Assert.Equal(WordComparisonReportWriter.ToMarkdown(result), result.ToMarkdown());
            Assert.Equal(WordComparisonReportWriter.ToTextSummary(result), result.ToTextSummary());

            using JsonDocument parsed = JsonDocument.Parse(result.ToJson());
            JsonElement root = parsed.RootElement;
            Assert.True(root.GetProperty("hasChanges").GetBoolean());
            Assert.Equal(result.Findings.Count, root.GetProperty("findingCount").GetInt32());
            Assert.True(root.GetProperty("summary").GetProperty("byScope").GetProperty("Paragraph").GetInt32() >= 1);
            Assert.Equal("Paragraph", root.GetProperty("findings")[0].GetProperty("scope").GetString());
            Assert.Equal("Modified", root.GetProperty("findings")[0].GetProperty("changeKind").GetString());
            Assert.Equal(root.GetProperty("findings")[0].GetProperty("location").GetString(), root.GetProperty("findings")[0].GetProperty("detailedLocation").GetString());
            Assert.Contains("Alpha | one", root.GetProperty("findings")[0].GetProperty("sourceText").GetString(), StringComparison.Ordinal);
            Assert.Contains("Beta | two", root.GetProperty("findings")[0].GetProperty("targetText").GetString(), StringComparison.Ordinal);

            string markdown = result.ToMarkdown();
            Assert.Contains("# Word Comparison Report", markdown, StringComparison.Ordinal);
            Assert.Contains("## By Scope", markdown, StringComparison.Ordinal);
            Assert.Contains("Detailed Location", markdown, StringComparison.Ordinal);
            Assert.Contains("| Paragraph |", markdown, StringComparison.Ordinal);
            Assert.Contains("Alpha \\| one", markdown, StringComparison.Ordinal);
            Assert.Contains("Beta \\| two", markdown, StringComparison.Ordinal);

            string summary = result.ToTextSummary();
            Assert.Contains("detected", summary, StringComparison.Ordinal);
            Assert.Contains("Paragraph=", summary, StringComparison.Ordinal);
            Assert.Contains("Modified=", summary, StringComparison.Ordinal);
        }

        [Fact]
        public void CompareStructureReportsNoChangesInSerializers() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_structure_report_same_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Same");
                document.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_structure_report_same_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Same");
                document.Save();
            }

            WordComparisonResult result = WordDocumentComparer.CompareStructure(sourcePath, targetPath);

            Assert.False(result.HasChanges);
            using JsonDocument parsed = JsonDocument.Parse(result.ToJson());
            Assert.False(parsed.RootElement.GetProperty("hasChanges").GetBoolean());
            Assert.Equal(0, parsed.RootElement.GetProperty("findingCount").GetInt32());
            Assert.Equal(JsonValueKind.Object, parsed.RootElement.GetProperty("summary").GetProperty("byScope").ValueKind);

            Assert.Contains("_None._", result.ToMarkdown(), StringComparison.Ordinal);
            Assert.Equal("No structural differences detected.", result.ToTextSummary());
        }

        [Fact]
        public void ComparisonReportWriterRejectsNullResult() {
            Assert.Throws<ArgumentNullException>(() => WordComparisonReportWriter.ToJson(null!));
            Assert.Throws<ArgumentNullException>(() => WordComparisonReportWriter.ToMarkdown(null!));
            Assert.Throws<ArgumentNullException>(() => WordComparisonReportWriter.ToTextSummary(null!));
        }
    }
}
