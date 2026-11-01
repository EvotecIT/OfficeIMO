using System.Management.Automation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;

namespace PSWriteOffice.Tests;

public sealed partial class ConversionCmdletWorkflowTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void GoogleExportModesResolveFileInputsAndBorrowedContentScopesWithoutChangingSources(bool excel) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-google-inputs-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        IDisposable? borrowed = null;
        try {
            string extension = excel ? ".xlsx" : ".docx";
            string literal = Path.Combine(root, "literal[1]" + extension), second = Path.Combine(root, "second" + extension);
            foreach (string path in new[] { literal, second }) {
                if (excel) {
                    using var source = ExcelDocument.Create();
                    source.AddSheet("Data").CellValue(1, 1, "Saved source"); source.Save(path);
                } else {
                    using var source = WordDocument.Create();
                    source.AddParagraph("Saved source"); source.Save(path);
                }
            }
            byte[] original = File.ReadAllBytes(literal);
            if (excel) {
                var document = ExcelDocument.Create();
                document.AddSheet("Data").CellValue(1, 1, "Borrowed source"); borrowed = document;
            } else {
                var document = WordDocument.Create(); document.AddParagraph("Borrowed source"); borrowed = document;
            }
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("document", borrowed);
            runspace.SessionStateProxy.SetVariable("excel", excel);
            runspace.SessionStateProxy.SetVariable("literal", literal);
            runspace.SessionStateProxy.SetVariable("second", second);
            runspace.SessionStateProxy.SetVariable("pattern", Path.Combine(root, "literal*" + extension));
            runspace.SessionStateProxy.SetVariable("scopedPath", Path.Combine(root, "scoped" + extension));
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddScript("""
                $command = if ($excel) { 'Export-OfficeExcelGoogleSpreadsheet' } else { 'Export-OfficeWordGoogleDocument' }
                & $command -LiteralPath $literal -Mode Batch -Options @{ Title = 'Literal report' } -ErrorAction Stop
                @([IO.FileInfo]::new($literal) | & $command -Mode Plan -ErrorAction Stop).Count
                @(& $command -Path @($pattern, $second) -Mode Plan -ErrorAction Stop).Count
                & $command -Document $document -Mode Batch -Options @{ Title = 'Borrowed report' } -ErrorAction Stop
                if ($excel) {
                    New-OfficeExcel -Path $scopedPath {
                        ExcelSheet Data { ExcelCell -Address A1 -Value 'Scoped source' }
                        $script:scopeBatch = Export-OfficeExcelGoogleSpreadsheet -Mode Batch -Options @{ Title = 'Scoped report' } -ErrorAction Stop
                    } | Out-Null
                } else {
                    New-OfficeWord -Path $scopedPath {
                        WordParagraph -Text 'Scoped source'
                        $script:scopeBatch = Export-OfficeWordGoogleDocument -Mode Batch -Options @{ Title = 'Scoped report' } -ErrorAction Stop
                    } | Out-Null
                }
                $scopeBatch
                """);
            var output = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(5, output.Count);
            void AssertBatch(PSObject value, string title, string text) {
                if (excel) {
                    var batch = Assert.IsType<GoogleSheetsBatch>(value.BaseObject);
                    Assert.Equal(title, batch.Title);
                    Assert.Contains(batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>().SelectMany(request => request.Cells),
                        cell => cell.Value.Value is string content && content.Contains(text, StringComparison.Ordinal));
                } else {
                    var batch = Assert.IsType<GoogleDocsBatch>(value.BaseObject);
                    Assert.Equal(title, batch.Title);
                    Assert.Contains(batch.Requests.OfType<GoogleDocsInsertParagraphRequest>(),
                        request => request.Paragraph.Text.Contains(text, StringComparison.Ordinal));
                }
            }
            AssertBatch(output[0], "Literal report", "Saved source");
            Assert.Equal(1, output[1].BaseObject); Assert.Equal(2, output[2].BaseObject);
            AssertBatch(output[3], "Borrowed report", "Borrowed source");
            AssertBatch(output[4], "Scoped report", "Scoped source");
            Assert.Equal(original, File.ReadAllBytes(literal));
            using (var exclusive = new FileStream(literal, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) Assert.True(exclusive.CanWrite);
            if (borrowed is ExcelDocument workbook) {
                Assert.True(workbook.Sheets[0].TryGetCellText(1, 1, out string text));
                Assert.Equal("Borrowed source", text);
            } else Assert.Equal("Borrowed source", ((WordDocument)borrowed).Paragraphs[0].Text);
        } finally { borrowed?.Dispose(); Directory.Delete(root, true); }
    }
}
