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
                source.AddWorksheet("Data").CellValue(1, 1, "Template content");
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

    [Fact]
    public void WordOpenContentPublishesOnlyAfterSuccessAndRejectsReadOnlyAuthoring() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-word-open-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "Review[final].docx");
            using (var source = WordDocument.Create()) {
                source.AddParagraph("Original");
                source.Save(path);
            }
            using var runspace = OpenWordRunspace();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            var output = shell.AddCommand("Open-OfficeWord").AddParameter("LiteralPath", path)
                .AddParameter("Content", ScriptBlock.Create("Add-OfficeWordParagraph -Text 'Saved in a block'")).Invoke();
            foreach (var item in output) if (item.BaseObject is WordDocument live) live.Dispose();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Empty(output);
            using (var saved = WordDocument.Load(path))
                Assert.Contains(saved.Paragraphs, paragraph => paragraph.Text == "Saved in a block");

            byte[] original = File.ReadAllBytes(path);
            shell.Commands.Clear();
            shell.AddCommand("Open-OfficeWord").AddParameter("LiteralPath", path)
                .AddParameter("Content", ScriptBlock.Create("Add-OfficeWordParagraph -Text 'Discard this'; throw 'block-failed'"));
            Exception? failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Single(Directory.GetFiles(directory));

            shell.Commands.Clear();
            shell.Streams.Error.Clear();
            runspace.SessionStateProxy.SetVariable("entered", false);
            shell.AddCommand("Open-OfficeWord").AddParameter("LiteralPath", path).AddParameter("ReadOnly")
                .AddParameter("Content", ScriptBlock.Create("$global:entered = $true"));
            failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
            Assert.False((bool)runspace.SessionStateProxy.GetVariable("entered"));
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public void WordFilePipelineSaveReturnsTheFileAndPreservesLoadedEncryptionUntilClose() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-word-encrypted-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "Encrypted[review].docx");
            string copy = Path.Combine(directory, "copy.docx");
            using (var source = WordDocument.Create()) {
                source.AddParagraph("Encrypted source");
                source.Save(path, new WordSaveOptions { Password = "test-secret" });
            }
            byte[] original = File.ReadAllBytes(path);
            using var runspace = OpenWordRunspace();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            using var document = Assert.IsType<WordDocument>(Assert.Single(shell.AddCommand("Open-OfficeWord")
                .AddParameter("Password", "test-secret").Invoke(new object[] { new FileInfo(path) })).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            document.AddParagraph("Saved through the pipeline");
            shell.Commands.Clear();
            var savedFile = Assert.Single(shell.AddCommand("Save-OfficeWord").AddParameter("Document", document)
                .AddParameter("Path", copy).AddParameter("PassThru").Invoke());
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(copy, Assert.IsType<FileInfo>(savedFile.BaseObject).FullName);
            Assert.False(document.IsDisposed);
            using (var saved = WordDocument.Load(copy, new WordLoadOptions { Password = "test-secret" }))
                Assert.Contains(saved.Paragraphs, paragraph => paragraph.Text == "Saved through the pipeline");
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }, File.ReadAllBytes(copy).Take(4));
            shell.Commands.Clear();
            shell.AddCommand("Close-OfficeWord").AddParameter("Document", document).Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.True(document.IsDisposed);
        } finally {
            Directory.Delete(directory, true);
        }
    }

    [Fact]
    public void WordCloseWhatIfCannotWriteThroughImplicitSaveOnDispose() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-word-close-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "auto-save.docx");
            using var document = WordDocument.Create(path, new WordCreateOptions {
                PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
            });
            document.Save(path);
            byte[] original = File.ReadAllBytes(path);
            document.AddParagraph("Do not publish through Close");
            using var runspace = OpenWordRunspace();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddCommand("Close-OfficeWord").AddParameter("Document", document).AddParameter("WhatIf");
            Exception? failure = Record.Exception(() => shell.Invoke());
            Assert.True(failure != null || shell.HadErrors);
            Assert.False(document.IsDisposed);
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally {
            Directory.Delete(directory, true);
        }
    }

    private static Runspace OpenWordRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("Open-OfficeWord", typeof(OpenOfficeWordCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Save-OfficeWord", typeof(SaveOfficeWordCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Close-OfficeWord", typeof(CloseOfficeWordCommand), null));
        state.Commands.Add(new SessionStateCmdletEntry("Add-OfficeWordParagraph", typeof(AddOfficeWordParagraphCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
