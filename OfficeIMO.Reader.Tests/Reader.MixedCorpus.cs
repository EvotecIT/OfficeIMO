using OfficeIMO.Email.Store.Tests;
using OfficeIMO.Excel;
using OfficeIMO.Reader;
using OfficeIMO.Word;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderMixedCorpusTests {
    [Fact]
    public async Task OneReaderCanDiscoverReadAndSearchWordExcelMarkdownPstAndOst() {
        const string query = "Synthetic";
        string root = Path.Combine(Path.GetTempPath(), "officeimo-reader-mixed-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);

        string wordPath = Path.Combine(root, "01-word.docx");
        string excelPath = Path.Combine(root, "02-excel.xlsx");
        string markdownPath = Path.Combine(root, "03-notes.md");
        string pstPath = Path.Combine(root, "04-mail.pst");
        string ostPath = Path.Combine(root, "05-offline.ost");
        string ignoredPath = Path.Combine(root, "ignored.bin");

        try {
            using (WordDocument document = WordDocument.Create(wordPath)) {
                document.AddParagraph("Synthetic contract in Word");
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Create(excelPath)) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.Cell(1, 1, "Synthetic contract in Excel");
                document.Save();
            }

            File.WriteAllText(markdownPath, "# Synthetic contract in Markdown");
            File.WriteAllBytes(pstPath, PstTestFileBuilder.Create());
            File.WriteAllBytes(ostPath, PstTestFileBuilder.Create(ost: true));
            File.WriteAllText(ignoredPath, "Synthetic unsupported input");

            OfficeDocumentReader reader = OfficeIMO.Reader.Tests.ReaderTestReaders.All;
            string[] paths = reader.EnumerateDocumentPaths(
                new[] { root },
                new ReaderFolderOptions {
                    Recurse = true,
                    MaxFiles = int.MaxValue
                }).ToArray();

            Assert.Equal(
                new[] { wordPath, excelPath, markdownPath, pstPath, ostPath },
                paths);

            IReadOnlyList<ReaderDocumentReadOutcome> outcomes = await reader.ReadDocumentsDetailedAsync(
                paths,
                batchOptions: new ReaderBatchOptions {
                    MaxDocuments = int.MaxValue,
                    MaxDegreeOfParallelism = 4
                });

            Assert.All(outcomes, outcome => Assert.True(outcome.Succeeded, outcome.Error?.ToString()));
            Assert.Equal(
                new[] {
                    ReaderInputKind.Word,
                    ReaderInputKind.Excel,
                    ReaderInputKind.Markdown,
                    ReaderInputKind.Email,
                    ReaderInputKind.Email
                },
                outcomes.Select(outcome => outcome.Document!.Kind).ToArray());

            OfficeDocumentSearchResult[] searches = outcomes
                .Select(outcome => outcome.Document!.Search(query, new OfficeDocumentSearchOptions {
                    MaximumResults = int.MaxValue
                }))
                .ToArray();

            Assert.All(searches, result => Assert.NotEmpty(result.Hits));
            Assert.Equal(paths, searches.Select(result => result.Source.Path).ToArray());
            Assert.All(searches, result => Assert.False(result.MaximumResultsReached));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }
}
