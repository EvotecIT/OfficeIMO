using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Excel;
using OfficeIMO.Word;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PSWriteOffice.Cmdlets.Excel;
using PSWriteOffice.Cmdlets.Word;
using PSWriteOffice.Cmdlets.OpenDocument;
using OfficeIMO.OpenDocument;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class DocumentLifecycleCmdletTests {
    [Theory]
    [InlineData("Text", ".odt", "application/vnd.oasis.opendocument.text")]
    [InlineData("Spreadsheet", ".ods", "application/vnd.oasis.opendocument.spreadsheet")]
    [InlineData("Presentation", ".odp", "application/vnd.oasis.opendocument.presentation")]
    [InlineData("Graphics", ".odg", "application/vnd.oasis.opendocument.graphics")]
    public void OpenDocumentCreationInfersItsNativeKindAndAssociatesThePathWithoutWriting(string kind, string extension, string mediaType) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-odf-lifecycle-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "Document" + extension);
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("New-OfficeOpenDocument", typeof(NewOfficeOpenDocumentCommand), null));
            state.Commands.Add(new SessionStateCmdletEntry("Save-OfficeOpenDocument", typeof(SaveOfficeOpenDocumentCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            var document = Assert.IsAssignableFrom<OdfDocument>(Assert.Single(shell.AddCommand("New-OfficeOpenDocument")
                .AddParameter("Path", path).Invoke()).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(kind, document.Kind.ToString());
            Assert.Equal(path, document.FilePath);
            Assert.Empty(Directory.GetFiles(root));
            shell.Commands.Clear();
            shell.AddCommand("Save-OfficeOpenDocument").AddParameter("Document", document).Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            using var package = System.IO.Compression.ZipFile.OpenRead(path);
            using var input = new StreamReader(package.GetEntry("mimetype")!.Open());
            Assert.Equal(mediaType, input.ReadToEnd());
        } finally {
            Directory.Delete(root, true);
        }
    }

    [Theory]
    [InlineData("Word", ".docx", ".docx")]
    [InlineData("Word", ".docx", ".dotx")]
    [InlineData("Excel", ".xlsx", ".xlsx")]
    [InlineData("Excel", ".xlsx", ".xltx")]
    public void LiveTemplatePreservesSourceAndDestinationUntilExplicitSave(string format, string extension, string templateExtension) {
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
            shell.AddCommand("New-Office" + format).AddParameter("Template", template)
                .AddParameter("Path", destination).AddParameter("Force");
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

    [Fact]
    public void ExcelSaveKeepsLoadedEncryptionAndReturnsAFileWithoutClosingTheWorkbook() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-excel-encrypted-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string path = Path.Combine(directory, "Encrypted[review].xlsx");
            string copy = Path.Combine(directory, "copy.xlsx");
            string plain = Path.Combine(directory, "plain.xlsx");
            using (var source = ExcelDocument.Create()) {
                source.AddSheet("Data").CellValue(1, 1, "Original");
                source.Save(path, new ExcelSaveOptions { Password = "test-secret" });
            }
            byte[] original = File.ReadAllBytes(path);
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Open-OfficeExcel", typeof(OpenOfficeExcelCommand), null));
            state.Commands.Add(new SessionStateCmdletEntry("New-OfficeExcelSaveOptions", typeof(NewOfficeExcelSaveOptionsCommand), null));
            state.Commands.Add(new SessionStateCmdletEntry("Save-OfficeExcel", typeof(SaveOfficeExcelCommand), null));
            state.Commands.Add(new SessionStateCmdletEntry("Close-OfficeExcel", typeof(CloseOfficeExcelCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state);
            runspace.Open();
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            using var document = Assert.IsType<ExcelDocument>(Assert.Single(shell.AddCommand("Get-Item")
                .AddParameter("LiteralPath", path).AddCommand("Open-OfficeExcel")
                .AddParameter("Password", "test-secret").AddParameter("MaxDegreeOfParallelism", 1).Invoke()).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(1, document.Execution.MaxDegreeOfParallelism);
            document.Sheets[0].CellValue(1, 1, "Reviewed");
            shell.Commands.Clear();
            var saved = Assert.Single(shell.AddCommand("Save-OfficeExcel").AddParameter("Document", document)
                .AddParameter("Path", copy).AddParameter("PassThru").Invoke());
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }, File.ReadAllBytes(copy).Take(4));
            Assert.Equal(copy, Assert.IsType<FileInfo>(saved.BaseObject).FullName);
            using (var loaded = ExcelDocument.Load(copy, new ExcelLoadOptions { Password = "test-secret" })) {
                Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string text));
                Assert.Equal("Reviewed", text);
            }
            Assert.Equal(original, File.ReadAllBytes(path));
            shell.Commands.Clear();
            var options = Assert.IsType<ExcelSaveOptions>(Assert.Single(shell.AddCommand("New-OfficeExcelSaveOptions")
                .AddParameter("Preflight").AddParameter("ValidateOpenXml").AddParameter("EvaluateFormulas").Invoke()).BaseObject);
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            document.Sheets[0].CellFormula(2, 2, "20+22");
            shell.Commands.Clear();
            shell.AddCommand("Save-OfficeExcel").AddParameter("Document", document).AddParameter("Path", plain)
                .AddParameter("Password", null).AddParameter("SaveOptions", options).Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Null(options.Password);
            Assert.True(options.SafePreflight);
            using (var loaded = ExcelDocument.Load(plain)) {
                Assert.True(loaded.Sheets[0].TryGetCellText(1, 1, out string text));
                Assert.Equal("Reviewed", text);
            }
            using (var package = SpreadsheetDocument.Open(plain, false)) {
                var formulaCell = package.WorkbookPart!.WorksheetParts.Single().Worksheet
                    .Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single(cell => cell.CellReference == "B2");
                Assert.Equal("42", formulaCell.CellValue!.Text);
            }
            string reusedPolicy = Path.Combine(directory, "reused-policy.xlsx");
            using (var reopened = ExcelDocument.Load(copy, new ExcelLoadOptions { Password = "test-secret" }))
                reopened.Save(reusedPolicy, options);
            Assert.Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }, File.ReadAllBytes(reusedPolicy).Take(4));
            string hashtableReset = Path.Combine(directory, "hashtable-reset.xlsx");
            using (var reopened = ExcelDocument.Load(copy, new ExcelLoadOptions { Password = "test-secret" })) {
                shell.Commands.Clear();
                shell.AddCommand("Save-OfficeExcel").AddParameter("Document", reopened).AddParameter("Path", hashtableReset)
                    .AddParameter("SaveOptions", new System.Collections.Hashtable { ["Password"] = null }).Invoke();
                Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            }
            Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, File.ReadAllBytes(hashtableReset).Take(4));
            using (var reopened = ExcelDocument.Load(hashtableReset)) {
                Assert.True(reopened.Sheets[0].TryGetCellText(1, 1, out string text));
                Assert.Equal("Reviewed", text);
            }
            byte[] savedPlain = File.ReadAllBytes(plain);
            document.Sheets[0].CellValue(1, 1, "Do not save this on close");
            shell.Commands.Clear();
            shell.AddCommand("Close-OfficeExcel").AddParameter("Document", document).Invoke();
            Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
            Assert.Equal(savedPlain, File.ReadAllBytes(plain));
            Assert.Equal(original, File.ReadAllBytes(path));

            string implicitPath = Path.Combine(directory, "implicit-save.xlsx");
            using var implicitDocument = ExcelDocument.Create(implicitPath, new ExcelCreateOptions {
                PersistenceMode = OfficeIMO.DocumentPersistenceMode.SaveOnDispose
            });
            implicitDocument.AddSheet("Data").CellValue(1, 1, "Original");
            implicitDocument.Save();
            byte[] beforeClose = File.ReadAllBytes(implicitPath);
            implicitDocument.Sheets[0].CellValue(1, 1, "Do not publish implicitly");
            shell.Commands.Clear();
            shell.AddCommand("Close-OfficeExcel").AddParameter("Document", implicitDocument);
            Exception? failure = Record.Exception(() => shell.Invoke());
            Assert.Equal(beforeClose, File.ReadAllBytes(implicitPath));
            Assert.True(failure != null || shell.HadErrors);
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
