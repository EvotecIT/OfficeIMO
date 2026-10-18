using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Pdf;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class DocumentPdfCmdletTests {
    [Fact]
    public async Task StoppingDuringPdfGenerationDoesNotPublishTheDestination() {
        string root = Path.Combine(AppContext.BaseDirectory, "pswriteoffice-pdf-stop-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        using var shaper = new PausedShaper();
        try {
            string font = new[] {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.ttf"),
                "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                "/Library/Fonts/Arial.ttf",
                "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
                "/System/Library/Fonts/Supplemental/Arial.ttf"
            }.First(File.Exists);
            var options = new PdfOptions { DefaultFont = PdfStandardFont.Courier, TextShapingProvider = shaper }
                .EmbedStandardFont(PdfStandardFont.Courier, font);
            var conversion = PdfPlainTextConverter.ToPdfDocumentResult("Stop during actual PDF generation", new PdfPlainTextOptions { PdfOptions = options });
            string output = Path.Combine(root, "stopped.pdf");
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Export-OfficeDocumentPdf", typeof(ExportOfficeDocumentPdfCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddCommand("Export-OfficeDocumentPdf").AddParameter("Document", conversion).AddParameter("Path", output);
            IAsyncResult invocation = shell.BeginInvoke();
            Assert.True(await Task.Run(() => shaper.Entered.Wait(TimeSpan.FromSeconds(15))), "The real font shaper was not reached during save.");
            IAsyncResult stopping = shell.BeginStop(null, null);
            Assert.True(SpinWait.SpinUntil(() => shell.InvocationStateInfo.State == PSInvocationState.Stopping, TimeSpan.FromSeconds(5)));
            shaper.Release.Set();
            await Task.Run(() => shell.EndStop(stopping));
            Assert.Throws<PipelineStoppedException>(() => shell.EndInvoke(invocation));
            Assert.False(File.Exists(output));
            Assert.Empty(Directory.GetFiles(root, "*.tmp"));
        } finally { shaper.Release.Set(); Directory.Delete(root, true); }
    }

    private sealed class PausedShaper : OfficeIMO.Drawing.IOfficeTextShapingProvider, IDisposable {
        public ManualResetEventSlim Entered { get; } = new();
        public ManualResetEventSlim Release { get; } = new();
        public OfficeIMO.Drawing.OfficeTextShapingResult? ShapeText(OfficeIMO.Drawing.OfficeTextShapingRequest request) {
            Entered.Set();
            if (!Release.Wait(TimeSpan.FromSeconds(20))) throw new TimeoutException("Font shaping cancellation test timed out.");
            return null;
        }
        public void Dispose() { Entered.Dispose(); Release.Dispose(); }
    }

    [Theory]
    [InlineData(".txt")]
    [InlineData(".doc")]
    public void FileExportProducesReopenablePdfAndWhatIfPreservesTheDestination(string extension) {
        string root = Path.Combine(AppContext.BaseDirectory, "pswriteoffice-pdf-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "source" + extension), output = Path.Combine(root, "output.pdf");
            if (extension == ".txt") File.WriteAllText(input, "  # Literal text\n<b>kept as text</b>");
            else { using var word = WordDocument.Create(); word.AddParagraph("Legacy DOC evidence"); word.Save(input); }
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Export-OfficeDocumentPdf", typeof(ExportOfficeDocumentPdfCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddCommand("Export-OfficeDocumentPdf").AddParameter("InputPath", input).AddParameter("Path", output).AddParameter("WhatIf");
            shell.Invoke(); Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error)); Assert.False(File.Exists(output));
            shell.Commands.Clear();
            byte[] original = File.ReadAllBytes(input);
            using (var rejecting = PowerShell.Create()) {
                rejecting.Runspace = runspace;
                rejecting.AddCommand("Export-OfficeDocumentPdf").AddParameter("InputPath", input).AddParameter("Path", input);
                Assert.Throws<CmdletInvocationException>(() => rejecting.Invoke());
                Assert.Equal(original, File.ReadAllBytes(input));
            }
            shell.AddCommand("Export-OfficeDocumentPdf").AddParameter("InputPath", input).AddParameter("Path", output).AddParameter("PassThru");
            Assert.Single(shell.Invoke()); Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error));
            Assert.Single(PdfDocument.Load(output).Read().Pages);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void DirectoryBatchResumesAndWhatIfCreatesNoState() {
        string root = Path.Combine(AppContext.BaseDirectory, "pswriteoffice-archive-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "source"), output = Path.Combine(root, "output"), checkpoint = Path.Combine(root, "checkpoint");
        Directory.CreateDirectory(input);
        try {
            File.WriteAllText(Path.Combine(input, "one.txt"), "Archive cmdlet evidence");
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Export-OfficeDocumentPdf", typeof(ExportOfficeDocumentPdfCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            void AddCommand() => shell.AddCommand("Export-OfficeDocumentPdf").AddParameter("InputDirectory", input)
                .AddParameter("OutputDirectory", output).AddParameter("CheckpointDirectory", checkpoint);
            AddCommand(); shell.AddParameter("WhatIf"); shell.Invoke();
            Assert.False(Directory.Exists(output)); Assert.False(Directory.Exists(checkpoint));
            shell.Commands.Clear(); AddCommand();
            Assert.Single(shell.Invoke()); Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error));
            shell.Commands.Clear(); AddCommand();
            var resumed = Assert.IsType<OfficeIMO.Workflows.OfficeConversionBatchResult>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.Equal(1, resumed.Reused);
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FailedBatchBoundsWarningsWithoutLosingSummaryOrStructuredOutcomes(bool itemResults) {
        string root = Path.Combine(AppContext.BaseDirectory, "pswriteoffice-failed-batch-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "source");
        Directory.CreateDirectory(input);
        try {
            string[] paths = Enumerable.Range(0, 103).Select(index => Path.Combine(input, index + ".docx")).ToArray();
            foreach (string path in paths) {
                File.WriteAllText(path, "Malformed package fixture");
            }
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Export-OfficeDocumentPdf", typeof(ExportOfficeDocumentPdfCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddCommand("Export-OfficeDocumentPdf").AddParameter("OutputDirectory", Path.Combine(root, "output"))
                .AddParameter("MaximumConcurrency", 4).AddParameter("SummaryVariable", "summary");
            if (itemResults) {
                shell.AddParameter("InputPaths", paths).AddParameter("ItemResults");
            } else {
                shell.AddParameter("InputDirectory", input);
            }
            var output = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error));
            var summary = Assert.IsType<OfficeIMO.Workflows.OfficeConversionBatchResult>(runspace.SessionStateProxy.GetVariable("summary"));
            Assert.Equal(103, summary.Selected); Assert.Equal(103, summary.Failed); Assert.Equal(0, summary.Completed);
            if (itemResults) {
                Assert.Equal(103, output.Count);
                Assert.All(output, item => Assert.Equal(OfficeIMO.Workflows.OfficeWorkflowStatus.Failed,
                    Assert.IsType<OfficeIMO.Workflows.OfficeConversionBatchItemResult>(item.BaseObject).Status));
                Assert.Empty(shell.Streams.Warning);
            } else {
                Assert.Same(summary, Assert.Single(output).BaseObject);
                Assert.Equal(101, shell.Streams.Warning.Count);
                Assert.Equal(100, shell.Streams.Warning.Count(warning => warning.Message.Contains(".docx:")));
                Assert.Equal("3 further failure messages were suppressed. Use -ItemResults for structured per-file outcomes.", shell.Streams.Warning[100].Message);
            }
            Assert.Empty(Directory.GetFiles(Path.Combine(root, "output")));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void BatchExportsAndResumesNativeEncryptedOutputWithoutSourceCredentials() {
        string root = Path.Combine(AppContext.BaseDirectory, "pswriteoffice-encrypted-batch-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "one.txt"), output = Path.Combine(root, "output"), checkpoint = Path.Combine(root, "state");
            File.WriteAllText(input, "Encrypted batch output");
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Export-OfficeDocumentPdf", typeof(ExportOfficeDocumentPdfCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            var settings = new PdfPlainTextOptions { PdfOptions = new PdfOptions().SetEncryption("Synthetic output reader", "Synthetic output owner") };
            void AddCommand() => shell.AddCommand("Export-OfficeDocumentPdf").AddParameter("InputPaths", new[] { input })
                .AddParameter("OutputDirectory", output).AddParameter("CheckpointDirectory", checkpoint).AddParameter("TextOptions", settings);
            AddCommand();
            var first = Assert.IsType<OfficeIMO.Workflows.OfficeConversionBatchResult>(Assert.Single(shell.Invoke()).BaseObject);
            Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error)); Assert.Equal(1, first.Completed); Assert.Equal(0, first.Failed);
            var encrypted = PdfDocument.Load(Path.Combine(output, "one.txt.pdf"),
                new PdfLoadOptions { Password = "Synthetic output reader" });
            Assert.True(encrypted.Inspect().Security.HasEncryption); Assert.Single(encrypted.Read().Pages);
            shell.Commands.Clear(); AddCommand();
            Assert.Equal(1, Assert.IsType<OfficeIMO.Workflows.OfficeConversionBatchResult>(Assert.Single(shell.Invoke()).BaseObject).Reused);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void FileInfoPipelineStreamsMixedOutcomesAndRetainsSummaryAndNativeOptions() {
        string root = Path.Combine(AppContext.BaseDirectory, "pswriteoffice-batch-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            File.WriteAllText(Path.Combine(root, "one.txt"), "<h1>literal pipeline text</h1>");
            File.WriteAllText(Path.Combine(root, "two.html"), "<h1>HTML pipeline text</h1>");
            File.WriteAllText(Path.Combine(root, "three.bin"), "Skipped input");
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Export-OfficeDocumentPdf", typeof(ExportOfficeDocumentPdfCommand), null));
            state.Commands.Add(new SessionStateCmdletEntry("New-OfficeHtmlPdfOptions", typeof(NewOfficeHtmlPdfOptionsCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            runspace.SessionStateProxy.SetVariable("source", root);
            runspace.SessionStateProxy.SetVariable("destination", Path.Combine(root, "output"));
            runspace.SessionStateProxy.SetVariable("textSettings", new PdfPlainTextOptions {
                PdfOptions = new PdfOptions { PageWidth = 240, PageHeight = 180, DefaultFontSize = 14 }
            });
            shell.AddScript("$htmlSettings = New-OfficeHtmlPdfOptions -Margin 12 -InteractiveFormControls $false; Get-ChildItem -LiteralPath $source -File | Export-OfficeDocumentPdf -OutputDirectory $destination -TextOptions $textSettings -HtmlOptions $htmlSettings -ItemResults -SummaryVariable summary");
            var results = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error));
            var items = results.Select(item => Assert.IsType<OfficeIMO.Workflows.OfficeConversionBatchItemResult>(item.BaseObject)).ToArray();
            Assert.Equal(3, items.Length);
            Assert.Single(items, item => item.Skipped);
            var summary = Assert.IsType<OfficeIMO.Workflows.OfficeConversionBatchResult>(runspace.SessionStateProxy.GetVariable("summary"));
            Assert.Equal(2, summary.Completed); Assert.Equal(1, summary.Skipped); Assert.Equal(0, summary.Failed);
            var textPdf = PdfDocument.Load(Path.Combine(root, "output", "one.txt.pdf")).Read();
            Assert.NotEmpty(textPdf.Pages);
            Assert.All(textPdf.Pages, page => Assert.Equal(240, page.Width));
            Assert.Contains("HTML pipeline text", PdfDocument.Load(Path.Combine(root, "output", "two.html.pdf")).Read().Text);
        } finally { Directory.Delete(root, true); }
    }
}
