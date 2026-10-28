using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Excel;
using OfficeIMO.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PSWriteOffice.Cmdlets.Excel;
using PSWriteOffice.Cmdlets.Word;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class DocumentLifecycleCmdletTests {
    [Theory]
    [InlineData("Word", ".docx", ".docx")]
    [InlineData("Word", ".docx", ".dotx")]
    [InlineData("Excel", ".xlsx", ".xlsx")]
    [InlineData("Excel", ".xlsx", ".xltx")]
    public void TemplateNoSavePreservesSourceAndDestinationUntilExplicitSave(string format, string extension, string templateExtension) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-template-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string template = Path.Combine(directory, "template" + templateExtension);
            string destination = Path.Combine(directory, "existing" + extension);
            if (format == "Word") {
                using var source = WordDocument.Create();
                source.AddParagraph("Template content");
                source.Save(template);
            } else {
                using var source = ExcelDocument.Create();
                source.AddSheet("Data").CellValue(1, 1, "Template content");
                source.Save(template);
            }
            byte[] sourceBytes = File.ReadAllBytes(template);
            byte[] original = { 1, 2, 3, 4 };
            File.WriteAllBytes(destination, original);
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("New-OfficeWord", typeof(NewOfficeWordCommand), null));
            state.Commands.Add(new SessionStateCmdletEntry("New-OfficeExcel", typeof(NewOfficeExcelCommand), null));
            state.Commands.Add(new SessionStateCmdletEntry("Save-OfficeWord", typeof(SaveOfficeWordCommand), null));
            state.Commands.Add(new SessionStateCmdletEntry("Save-OfficeExcel", typeof(SaveOfficeExcelCommand), null));
            state.Commands.Add(new SessionStateCmdletEntry("Close-OfficeWord", typeof(CloseOfficeWordCommand), null));
            state.Commands.Add(new SessionStateCmdletEntry("Close-OfficeExcel", typeof(CloseOfficeExcelCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("New-Office" + format).AddParameter("TemplatePath", template)
                .AddParameter("Path", destination).AddParameter("NoSave");
            var document = Assert.Single(shell.Invoke()).BaseObject;
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            try {
                Assert.Equal(original, File.ReadAllBytes(destination));
                Assert.Equal(sourceBytes, File.ReadAllBytes(template));
                shell.Commands.Clear();
                shell.AddCommand("Save-Office" + format).AddParameter("Document", document).Invoke();
                Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
                Assert.Equal(sourceBytes, File.ReadAllBytes(template));
                if (format == "Word") {
                    using (var package = WordprocessingDocument.Open(destination, false))
                        Assert.Equal(WordprocessingDocumentType.Document, package.DocumentType);
                    using var saved = WordDocument.Load(destination);
                    Assert.Contains(saved.Paragraphs, paragraph => paragraph.Text == "Template content");
                } else {
                    using (var package = SpreadsheetDocument.Open(destination, false))
                        Assert.Equal(SpreadsheetDocumentType.Workbook, package.DocumentType);
                    using var saved = ExcelDocument.Load(destination);
                    Assert.True(saved.Sheets[0].TryGetCellText(1, 1, out string text));
                    Assert.Equal("Template content", text);
                }
            } finally {
                shell.Commands.Clear();
                shell.AddCommand("Close-Office" + format).AddParameter("Document", document).Invoke();
            }
        } finally {
            Directory.Delete(directory, true);
        }
    }
}
