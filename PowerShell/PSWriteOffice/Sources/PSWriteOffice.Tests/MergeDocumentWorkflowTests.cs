using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Conversion;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class MergeDocumentWorkflowTests {
    [Fact]
    public void DirectoryAliasesCannotTurnMergePublicationIntoSourceReplacement() {
        using var files = new MergeFiles();
        string targetDirectory = files.PathFor("real");
        Directory.CreateDirectory(targetDirectory);
        string source = files.Word("real/source.docx", "SOURCE");
        string aliasDirectory = files.PathFor("alias");
        string output = Path.Combine(aliasDirectory, "source.docx");
        using var runspace = OpenRunspace();
        using var directoryAlias = PhysicalPathTestSupport.CreateDirectoryAlias(runspace, aliasDirectory, targetDirectory);
        var before = File.ReadAllBytes(source);
        Set(runspace, ("source", source), ("output", output));
        Assert.Contains("differ", Fail(runspace,
            "Merge-OfficeDocument -LiteralPath $source -OutputPath $output -ErrorAction Stop"));
        Assert.Equal(before, File.ReadAllBytes(source));
    }

    [Theory]
    [InlineData(".docx", true)]
    [InlineData(".xlsx", true)]
    [InlineData(".docx", false)]
    [InlineData(".xlsx", false)]
    public void SignedMergeTargetsAreRejectedBeforeLiveContentChangesOrPublication(string extension, bool live) {
        using var files = new MergeFiles();
        string target = extension == ".docx" ? files.Word("signed.docx", "TARGET") : files.Excel("signed.xlsx", "Target", "TARGET");
        string incoming = extension == ".docx" ? files.Word("incoming.docx", "INCOMING") : files.Excel("incoming.xlsx", "Incoming", "INCOMING");
        if (extension == ".docx") {
            using var package = WordprocessingDocument.Open(target, true);
            package.AddDigitalSignatureOriginPart();
        } else {
            using var package = SpreadsheetDocument.Open(target, true);
            package.AddDigitalSignatureOriginPart();
        }
        var before = files.Snapshot();
        string output = files.PathFor("absent/merged" + extension);
        using IDisposable document = extension == ".docx"
            ? WordDocument.Load(target, new WordLoadOptions { PersistenceMode = DocumentPersistenceMode.Explicit })
            : ExcelDocument.Load(target, new ExcelLoadOptions { PersistenceMode = DocumentPersistenceMode.Explicit });
        string[] state = document is WordDocument word
            ? word.Paragraphs.Select(paragraph => paragraph.Text).ToArray()
            : ((ExcelDocument)document).Sheets.Select(sheet => sheet.Name).ToArray();
        using var runspace = OpenRunspace();
        Set(runspace, ("document", document), ("inputs", live ? new[] { incoming } : new[] { target, incoming }), ("output", output));
        string arguments = live ? "-Document $document -Path $inputs" : "-LiteralPath $inputs";
        Assert.Contains("signed", Fail(runspace,
            "Merge-OfficeDocument " + arguments + " -OutputPath $output -ErrorAction Stop"), StringComparison.OrdinalIgnoreCase);
        Assert.Equal(state, document is WordDocument unchangedWord
            ? unchangedWord.Paragraphs.Select(paragraph => paragraph.Text).ToArray()
            : ((ExcelDocument)document).Sheets.Select(sheet => sheet.Name).ToArray());
        Assert.False(Directory.Exists(Path.GetDirectoryName(output)));
        files.AssertSourcesUnchanged(before);
    }

    [Theory]
    [InlineData(".docx", WordDocumentType.Document)]
    [InlineData(".docm", WordDocumentType.MacroEnabledDocument)]
    [InlineData(".dotx", WordDocumentType.Template)]
    [InlineData(".dotm", WordDocumentType.MacroEnabledTemplate)]
    public void WordStringPipelinePreservesOrderSourcesAndPackageKind(string extension, WordDocumentType kind) {
        using var files = new MergeFiles();
        string first = files.Word("first" + extension, "FIRST", kind);
        string second = files.Word("second" + extension, "SECOND", kind);
        string third = files.Word("third" + extension, "THIRD", kind);
        string output = files.PathFor("merged" + extension);
        var sources = files.Snapshot();
        using var runspace = OpenRunspace();
        Set(runspace, ("inputs", new[] { third, first, second }), ("output", output));
        var result = Invoke(runspace, "$inputs | Merge-OfficeDocument -OutputPath $output -PassThru -Options @{ SectionMode = 'UseDestinationSections' } -ErrorAction Stop");
        Assert.Equal(output, Assert.IsType<FileInfo>(Assert.Single(result).BaseObject).FullName);
        using var saved = WordDocument.Load(output);
        Assert.Equal(extension, saved.SourceFormatDescriptor.Extension);
        InOrder(string.Join("\n", saved.Paragraphs.Select(p => p.Text)), "THIRD", "FIRST", "SECOND");
        files.AssertSourcesUnchanged(sources);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FileInfoPipelineUsesLiteralFullNamesOutsideTheCurrentDirectory(bool pdf) {
        using var files = new MergeFiles();
        string first = pdf ? files.Pdf("source[1].pdf", "FIRST") : files.Word("part[2].docx", "FIRST");
        string second = pdf ? files.Pdf("source[2].pdf", "SECOND") : files.Word("part[1].docx", "SECOND");
        string output = files.PathFor(pdf ? "merged.pdf" : "merged.docx");
        var sources = files.Snapshot();
        using var runspace = OpenRunspace();
        Set(runspace, ("inputs", new[] { first, second }), ("output", output));
        Assert.Empty(Invoke(runspace, "Get-Item -LiteralPath $inputs | Merge-OfficeDocument -OutputPath $output -ErrorAction Stop"));
        if (pdf) {
            using var saved = PdfDocument.Load(output);
            Assert.Equal(2, saved.Inspect().Pages.Count);
            InOrder(saved.Read().Text, "FIRST", "SECOND");
        } else {
            using var saved = WordDocument.Load(output);
            InOrder(string.Join("\n", saved.Paragraphs.Select(p => p.Text)), "FIRST", "SECOND");
        }
        files.AssertSourcesUnchanged(sources);
    }

    [Fact]
    public void WordLivePipelineAppendsWithoutSavingOrTakingOwnership() {
        using var files = new MergeFiles();
        string first = files.Word("first.docx", "FIRST"), second = files.Word("second.docx", "SECOND");
        using var live = WordDocument.Create();
        live.AddParagraph("COVER");
        var before = files.Snapshot();
        using var runspace = OpenRunspace();
        Set(runspace, ("live", live), ("inputs", new[] { first, second }));
        var result = Invoke(runspace, "$live | Merge-OfficeDocument -Path $inputs -PassThru -ErrorAction Stop");
        Assert.Same(live, Assert.Single(result).BaseObject);
        files.AssertUnchanged(before);
        string output = files.PathFor("live.docx");
        live.AddParagraph("STILL OPEN");
        live.Save(output);
        using var saved = WordDocument.Load(output);
        InOrder(string.Join("\n", saved.Paragraphs.Select(p => p.Text)), "COVER", "FIRST", "SECOND", "STILL OPEN");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExcelFileAndLiveMergesUseNativeOptionsAndPreserveSheetOrder(bool liveInput) {
        using var files = new MergeFiles();
        string first = files.Excel("first.xlsx", "First", "FIRST");
        string second = files.Excel("second.xlsx", "Second", "SECOND");
        string output = files.PathFor("merged.xlsx");
        using var live = ExcelDocument.Create();
        live.AddSheet("Cover").CellValue(1, 1, "COVER");
        var sources = files.Snapshot();
        var settings = new ExcelWorkbookMergeOptions { CopyMode = ExcelWorksheetCopyMode.Package };
        using var runspace = OpenRunspace();
        Set(runspace, ("live", live), ("inputs", new[] { first, second }), ("output", output), ("settings", settings));
        var result = Invoke(runspace, liveInput
            ? "$live | Merge-OfficeDocument -Path $inputs -OutputPath $output -Options $settings, @{ SheetNamePrefix = 'Imported ' } -PassThru -ErrorAction Stop"
            : "Merge-OfficeDocument -LiteralPath $inputs -OutputPath $output -Options @{ CopyMode = 'Package'; SheetNamePrefix = 'Imported ' } -PassThru -ErrorAction Stop");
        if (liveInput) Assert.Same(live, Assert.Single(result).BaseObject);
        else Assert.Equal(output, Assert.IsType<FileInfo>(Assert.Single(result).BaseObject).FullName);
        using var saved = ExcelDocument.Load(output);
        Assert.Equal(liveInput ? new[] { "Cover", "Imported First", "Imported Second" } : new[] { "First", "Imported Second" },
            saved.Sheets.Select(sheet => sheet.Name));
        var markers = liveInput ? new[] { "COVER", "FIRST", "SECOND" } : new[] { "FIRST", "SECOND" };
        for (int index = 0; index < saved.Sheets.Count; index++) {
            Assert.True(saved.Sheets[index].TryGetCellText(1, 1, out string text));
            Assert.Equal(markers[index], text);
            if (!liveInput || index > 0) Assert.Equal("0.0000", saved.Sheets[index].CellAt(2, 1).GetStyle().NumberFormatCode);
        }
        Assert.Null(settings.SheetNamePrefix);
        live.AddSheet("Still open");
        files.AssertSourcesUnchanged(sources);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PdfMergeAuthenticatesSingletonOrPerSourcePasswordsAndPreservesPageOrder(bool perSource) {
        using var files = new MergeFiles();
        string first = files.Pdf("first.pdf", "FIRST", "one");
        string second = files.Pdf("second.pdf", "SECOND", perSource ? "two" : "one");
        string output = files.PathFor("merged.pdf");
        var sources = files.Snapshot();
        using var one = Password("one");
        using var two = Password("two");
        using var runspace = OpenRunspace();
        var options = new PdfMergeOptions { FlattenVisualAnnotations = true };
        Set(runspace, ("inputs", new[] { second, first }), ("output", output), ("settings", options),
            ("passwords", perSource ? new[] { two, one } : new[] { one }));
        var result = Invoke(runspace, "Merge-OfficeDocument -LiteralPath $inputs -OutputPath $output -Password $passwords -Options $settings, @{ FlattenVisualAnnotations = $false } -PassThru -ErrorAction Stop");
        Assert.Equal(output, Assert.IsType<FileInfo>(Assert.Single(result).BaseObject).FullName);
        using var saved = PdfDocument.Load(output, new PdfLoadOptions { Password = perSource ? "two" : "one" });
        Assert.Equal(2, saved.Inspect().Pages.Count);
        InOrder(saved.Read().Text, "SECOND", "FIRST");
        Assert.True(options.FlattenVisualAnnotations);
        string unsupported = files.PathFor("unsupported/flattened.pdf");
        Set(runspace, ("unsupported", unsupported));
        string error = Fail(runspace, "Merge-OfficeDocument -LiteralPath $inputs -OutputPath $unsupported -Password $passwords -Options @{ FlattenVisualAnnotations = $true } -ErrorAction Stop");
        Assert.Contains("blocked", error, StringComparison.OrdinalIgnoreCase);
        Assert.False(Directory.Exists(Path.GetDirectoryName(unsupported)));
        files.AssertSourcesUnchanged(sources);
    }

    [Fact]
    public void PdfMergeAppliesNativePageNormalizationForUnencryptedInputs() {
        using var files = new MergeFiles();
        string first = files.Pdf("first.pdf", "FIRST");
        string second = files.Pdf("second.pdf", "SECOND");
        string output = files.PathFor("normalized.pdf");
        var sources = files.Snapshot();
        var options = new PdfMergeOptions {
            ResizePages = new PdfPageResizeOptions(PdfPageSizes.A4),
            Policy = new PdfMergePolicy { Outlines = PdfMergeStructureMode.Combine }
        };
        using var runspace = OpenRunspace();
        Set(runspace, ("inputs", new[] { first, second }), ("output", output), ("settings", options));
        Invoke(runspace, "Merge-OfficeDocument -LiteralPath $inputs -OutputPath $output -Options $settings, @{ FlattenVisualAnnotations = $true; Policy = @{ Metadata = 'Drop' } } -ErrorAction Stop");
        Assert.Equal(PdfMergeStructureMode.KeepPrimary, options.Policy.Metadata);
        Assert.Equal(PdfMergeStructureMode.Combine, options.Policy.Outlines);
        using var saved = PdfDocument.Load(output);
        Assert.Equal(2, saved.Inspect().Pages.Count);
        InOrder(saved.Read().Text, "FIRST", "SECOND");
        foreach (var page in saved.Inspect().Pages) {
            Assert.InRange(page.Width, 595, 596);
            Assert.InRange(page.Height, 841, 843);
        }
        Assert.False(options.FlattenVisualAnnotations);
        files.AssertSourcesUnchanged(sources);
    }

    [Fact]
    public void PdfMergeKeepsNativeSourcePermissionEnforcementAndDisposesFailedInputs() {
        using var files = new MergeFiles();
        string first = files.Pdf("first.pdf", "FIRST");
        string restricted = files.Pdf("restricted.pdf", "RESTRICTED", "user", PdfStandardPermissions.None);
        string output = files.PathFor("new/merged.pdf");
        var before = files.Snapshot();
        using var password = Password("user");
        using var runspace = OpenRunspace();
        Set(runspace, ("inputs", new[] { first, restricted }), ("output", output), ("password", password));
        string error = Fail(runspace, "Merge-OfficeDocument -LiteralPath $inputs -OutputPath $output -Password $password -ErrorAction Stop");
        Assert.Contains("permission", error, StringComparison.OrdinalIgnoreCase);
        files.AssertUnchanged(before);
        foreach (string path in new[] { first, restricted }) {
            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            Assert.True(exclusive.Length > 0);
        }
    }

    [Theory]
    [InlineData("-Options @{ UnknownOption = 1 }", "Unknown")]
    [InlineData("-Options @{ SectionMode = 99 }", "SectionMode")]
    [InlineData("-Options $wrongOptions", "WordDocumentMergeOptions")]
    [InlineData("-Password $tooMany", "Password")]
    public void BindingFailuresLeaveLiveTargetBytesAndDirectoriesUnchanged(string arguments, string expected) {
        using var files = new MergeFiles();
        string first = files.Word("first.docx", "FIRST"), second = files.Word("second.docx", "SECOND");
        string output = files.PathFor("new/merged.docx");
        using var live = WordDocument.Create();
        live.AddParagraph("COVER");
        using var password = Password("one");
        var before = files.Snapshot();
        using var runspace = OpenRunspace();
        Set(runspace, ("live", live), ("inputs", new[] { first, second }), ("output", output),
            ("wrongOptions", new PdfMergeOptions()), ("tooMany", new[] { password, password, password }));
        Assert.Contains(expected, Fail(runspace, "Merge-OfficeDocument -Document $live -Path $inputs -OutputPath $output " + arguments + " -ErrorAction Stop"));
        Assert.Equal(new[] { "COVER" }, live.Paragraphs.Select(p => p.Text));
        files.AssertUnchanged(before);
    }

    [Fact]
    public void WhatIfAndNoClobberPreserveExistingBytesAndDoNotCreateDirectoriesOrMutateLiveDocuments() {
        using var files = new MergeFiles();
        string source = files.Word("source.docx", "SOURCE");
        string existing = files.Word("existing.docx", "EXISTING");
        string missing = files.PathFor("absent/merged.docx");
        var before = files.Snapshot();
        using var live = WordDocument.Create();
        live.AddParagraph("COVER");
        using var runspace = OpenRunspace();
        Set(runspace, ("source", source), ("existing", existing), ("missing", missing), ("live", live));
        Assert.Empty(Invoke(runspace, "Merge-OfficeDocument -LiteralPath $source -OutputPath $missing -WhatIf -PassThru -Show -ErrorAction Stop"));
        Assert.Empty(Invoke(runspace, "Merge-OfficeDocument -LiteralPath $source -OutputPath $existing -WhatIf -ErrorAction Stop"));
        Assert.Empty(Invoke(runspace, "Merge-OfficeDocument -Document $live -Path $source -OutputPath $missing -WhatIf -PassThru -ErrorAction Stop"));
        Assert.Contains("already exists", Fail(runspace, "Merge-OfficeDocument -Document $live -Path $source -OutputPath $existing -NoClobber -ErrorAction Stop"));
        Assert.Contains("already exists", Fail(runspace, "Merge-OfficeDocument -LiteralPath $source -OutputPath $existing -NoClobber -ErrorAction Stop"));
        Assert.Equal(new[] { "COVER" }, live.Paragraphs.Select(p => p.Text));
        files.AssertUnchanged(before);
        // Default overwrite publishes a valid artifact while preserving the input.
        Assert.Empty(Invoke(runspace, "Merge-OfficeDocument -LiteralPath $source -OutputPath $existing -ErrorAction Stop"));
        using var saved = WordDocument.Load(existing);
        Assert.Contains("SOURCE", string.Join("\n", saved.Paragraphs.Select(p => p.Text)));
        Assert.Equal(before[source], File.ReadAllBytes(source));
    }

    [Fact]
    public void MixedUnsupportedMislabeledAndMismatchedInputsAreRejectedBeforeMutationOrPublication() {
        using var files = new MergeFiles();
        string word = files.Word("word.docx", "WORD"), excel = files.Excel("book.xlsx", "Data", "EXCEL");
        string template = files.Word("template.docx", "TEMPLATE", WordDocumentType.Template);
        string unsupported = files.PathFor("deck.pptx");
        File.WriteAllText(unsupported, "unsupported native family");
        string output = files.PathFor("new/merged.docx"), wrong = files.PathFor("new/merged.pdf");
        using var live = WordDocument.Create();
        live.AddParagraph("COVER");
        using var pdf = PdfDocument.Create();
        var before = files.Snapshot();
        using var runspace = OpenRunspace();
        Set(runspace, ("word", word), ("excel", excel), ("template", template), ("unsupported", unsupported),
            ("output", output), ("wrong", wrong), ("live", live), ("pdf", pdf));
        Assert.Contains("same format", Fail(runspace, "Merge-OfficeDocument -Path $word, $excel -OutputPath $output -ErrorAction Stop"));
        Assert.Contains("same format", Fail(runspace, "Merge-OfficeDocument -Document $live -Path $word -OutputPath $wrong -ErrorAction Stop"));
        Assert.Contains("Unsupported format", Fail(runspace, "Merge-OfficeDocument -Path $unsupported -OutputPath $output -ErrorAction Stop"));
        Assert.Contains("format", Fail(runspace, "Merge-OfficeDocument -Document $live -Path $word, $template -OutputPath $output -ErrorAction Stop"));
        Assert.Contains("differ", Fail(runspace, "Merge-OfficeDocument -Path $word -OutputPath $word -ErrorAction Stop"));
        Assert.Contains("Live PDF", Fail(runspace, "Merge-OfficeDocument -Document $pdf -Path $word -ErrorAction Stop"));
        Assert.Equal(new[] { "COVER" }, live.Paragraphs.Select(p => p.Text));
        files.AssertUnchanged(before);
    }

    [Fact]
    public void LaterBadPasswordDoesNotPartiallyAppendToLiveWordAndSourcesCanBeReopenedExclusively() {
        using var files = new MergeFiles();
        string first = files.Word("first.docx", "FIRST");
        string encrypted = files.PathFor("encrypted.docx");
        using (var protectedDocument = WordDocument.Create()) {
            protectedDocument.AddParagraph("PROTECTED");
            protectedDocument.Save(encrypted, new WordSaveOptions { Password = "correct" });
        }
        using var live = WordDocument.Create();
        live.AddParagraph("COVER");
        using var password = Password("incorrect");
        var before = files.Snapshot();
        using var runspace = OpenRunspace();
        Set(runspace, ("live", live), ("inputs", new[] { first, encrypted }), ("passwords", new SecureString[] { null!, password }));
        Assert.NotEmpty(Fail(runspace, "Merge-OfficeDocument -Document $live -Path $inputs -Password $passwords -ErrorAction Stop"));
        Assert.Equal(new[] { "COVER" }, live.Paragraphs.Select(p => p.Text));
        files.AssertUnchanged(before);
        foreach (string path in new[] { first, encrypted }) {
            using var exclusive = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void EncryptedOfficeInputsUseNativePasswordLoadingAndKeepPrimaryEncryption(bool excel, bool perSource) {
        using var files = new MergeFiles();
        string extension = excel ? ".xlsx" : ".docx";
        string first = files.PathFor("first" + extension), second = files.PathFor("second" + extension);
        string output = files.PathFor("merged" + extension);
        for (int index = 0; index < 2; index++) {
            string path = index == 0 ? first : second;
            string password = index == 1 && perSource ? "two" : "one";
            string marker = index == 0 ? "FIRST" : "SECOND";
            if (excel) {
                using var workbook = ExcelDocument.Create();
                workbook.AddSheet(marker).CellValue(1, 1, marker);
                workbook.Save(path, new ExcelSaveOptions { Password = password });
            } else {
                using var document = WordDocument.Create();
                document.AddParagraph(marker);
                document.Save(path, new WordSaveOptions { Password = password });
            }
        }
        var before = files.Snapshot();
        using var one = Password("one");
        using var two = Password("two");
        using var runspace = OpenRunspace();
        Set(runspace, ("inputs", new[] { first, second }), ("output", output),
            ("passwords", perSource ? new[] { one, two } : new[] { one }));
        Assert.Empty(Invoke(runspace, "Merge-OfficeDocument -LiteralPath $inputs -OutputPath $output -Password $passwords -ErrorAction Stop"));
        if (excel) {
            using var saved = ExcelDocument.Load(output, new ExcelLoadOptions { Password = "one" });
            Assert.Equal(new[] { "FIRST", "SECOND" }, saved.Sheets.Select(sheet => sheet.Name));
            for (int index = 0; index < 2; index++) {
                Assert.True(saved.Sheets[index].TryGetCellText(1, 1, out string text));
                Assert.Equal(index == 0 ? "FIRST" : "SECOND", text);
            }
        } else {
            using var saved = WordDocument.Load(output, new WordLoadOptions { Password = "one" });
            InOrder(string.Join("\n", saved.Paragraphs.Select(p => p.Text)), "FIRST", "SECOND");
        }
        files.AssertSourcesUnchanged(before);
    }

    private static Runspace OpenRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(MergeOfficeDocumentCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }

    private static void Set(Runspace runspace, params (string Name, object Value)[] variables) {
        foreach (var variable in variables) runspace.SessionStateProxy.SetVariable(variable.Name, variable.Value);
    }

    private static PSObject[] Invoke(Runspace runspace, string script) {
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript(script);
        var result = shell.Invoke();
        Assert.False(shell.HadErrors, string.Join(Environment.NewLine, shell.Streams.Error));
        return result.ToArray();
    }

    private static string Fail(Runspace runspace, string script) {
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddScript("try { " + script + "; throw 'MERGE_UNEXPECTED_SUCCESS' } catch { $_.Exception.Message }");
        var result = shell.Invoke();
        string message = Assert.IsType<string>(Assert.Single(result).BaseObject);
        Assert.DoesNotContain("MERGE_UNEXPECTED_SUCCESS", message);
        return message;
    }

    private static SecureString Password(string text) {
        var secure = new SecureString();
        foreach (char character in text) secure.AppendChar(character);
        secure.MakeReadOnly();
        return secure;
    }

    private static void InOrder(string text, params string[] markers) {
        int previous = -1;
        foreach (string marker in markers) {
            int next = text.IndexOf(marker, previous + 1, StringComparison.Ordinal);
            Assert.True(next > previous, "Missing or reordered marker: " + marker);
            previous = next;
        }
    }

    private sealed class MergeFiles : IDisposable {
        private readonly string _root = Directory.CreateTempSubdirectory("officeimo-merge-command-").FullName;
        internal string PathFor(string name) => Path.GetFullPath(Path.Combine(_root, name));
        internal string Word(string name, string text, WordDocumentType kind = WordDocumentType.Document) {
            string path = PathFor(name);
            using var document = WordDocument.Create(new WordCreateOptions { DocumentType = kind });
            document.AddParagraph(text);
            // Stream output preserves the actual package kind even when testing a mislabeled input.
            using var stream = File.Create(path);
            document.Save(stream);
            return path;
        }
        internal string Excel(string name, string sheet, string text) {
            string path = PathFor(name);
            using var document = ExcelDocument.Create();
            var worksheet = document.AddSheet(sheet);
            worksheet.CellValue(1, 1, text);
            worksheet.CellValue(2, 1, 0.25);
            worksheet.FormatCell(2, 1, "0.0000");
            document.Save(path);
            return path;
        }
        internal string Pdf(string name, string text, string? password = null,
            PdfStandardPermissions permissions = PdfStandardPermissions.All) {
            string path = PathFor(name);
            var options = new PdfCreateOptions();
            if (password != null) options.SetEncryption(new PdfStandardEncryptionOptions(password) {
                OwnerPassword = "owner", AllowedPermissions = permissions
            });
            using var document = PdfDocument.Create(options);
            document.Content.Paragraph(p => p.Text(text));
            document.Save(path);
            return path;
        }
        internal Dictionary<string, byte[]> Snapshot() => Directory.GetFiles(_root, "*", SearchOption.AllDirectories)
            .ToDictionary(path => path, File.ReadAllBytes);
        internal void AssertSourcesUnchanged(Dictionary<string, byte[]> before) {
            foreach (var entry in before) Assert.Equal(entry.Value, File.ReadAllBytes(entry.Key));
        }
        internal void AssertUnchanged(Dictionary<string, byte[]> before) {
            Assert.Equal(before.Keys.OrderBy(path => path), Snapshot().Keys.OrderBy(path => path));
            Assert.Empty(Directory.GetDirectories(_root, "*", SearchOption.AllDirectories));
            AssertSourcesUnchanged(before);
        }
        public void Dispose() => Directory.Delete(_root, true);
    }
}
