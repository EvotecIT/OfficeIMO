using OfficeIMO;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using PSWriteOffice.Cmdlets.Pdf;
using PSWriteOffice.Cmdlets.Conversion;

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
            var options = new PdfCreateOptions { DefaultFont = PdfStandardFont.Courier, TextShapingProvider = shaper }
                .EmbedStandardFont(PdfStandardFont.Courier, font);
            string input = Path.Combine(root, "input.txt");
            File.WriteAllText(input, "Stop during actual PDF generation");
            string output = Path.Combine(root, "stopped.pdf");
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Convert-OfficeDocument", typeof(ConvertOfficeDocumentCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddCommand("Convert-OfficeDocument").AddParameter("To", "Pdf").AddParameter("Path", input).AddParameter("OutputPath", output)
                .AddParameter("Options", new PdfTextImportOptions { PdfOptions = options });
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

    private static OfficeBatchSummary InvokeSummary(PowerShell shell, Runspace runspace) {
        shell.AddParameter("SummaryVariable", "summary");
        shell.Invoke();
        Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error));
        return Assert.IsType<OfficeBatchSummary>(runspace.SessionStateProxy.GetVariable("summary"));
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
            state.Commands.Add(new SessionStateCmdletEntry("Convert-OfficeDocument", typeof(ConvertOfficeDocumentCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddCommand("Convert-OfficeDocument").AddParameter("To", "Pdf").AddParameter("Path", input).AddParameter("OutputPath", output).AddParameter("WhatIf");
            shell.Invoke(); Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error)); Assert.False(File.Exists(output));
            shell.Commands.Clear();
            byte[] original = File.ReadAllBytes(input);
            using (var rejecting = PowerShell.Create()) {
                rejecting.Runspace = runspace;
                rejecting.AddCommand("Convert-OfficeDocument").AddParameter("To", "Pdf").AddParameter("Path", input).AddParameter("OutputPath", input);
                Assert.Throws<CmdletInvocationException>(() => rejecting.Invoke());
                Assert.Equal(original, File.ReadAllBytes(input));
            }
            shell.AddCommand("Convert-OfficeDocument").AddParameter("To", "Pdf").AddParameter("Path", input).AddParameter("OutputPath", output);
            Assert.Single(shell.Invoke()); Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error));
            Assert.Single(PdfDocument.Load(output).Read().Pages);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void FileExportReplacesAnExistingPdfUnlessNoClobberIsGiven() {
        string root = Path.Combine(AppContext.BaseDirectory, "pswriteoffice-pdf-force-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string input = Path.Combine(root, "source.txt"), output = Path.Combine(root, "output.pdf");
            File.WriteAllText(input, "Replacement export");
            byte[] existing = { 1, 2, 3 };
            File.WriteAllBytes(output, existing);
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Convert-OfficeDocument", typeof(ConvertOfficeDocumentCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using (var refusing = PowerShell.Create()) {
                refusing.Runspace = runspace;
                refusing.AddCommand("Convert-OfficeDocument").AddParameter("To", "Pdf").AddParameter("Path", input).AddParameter("OutputPath", output).AddParameter("NoClobber");
                var skipped = Assert.IsType<OfficeBatchItemReport>(Assert.Single(refusing.Invoke()).BaseObject);
                Assert.False(refusing.HadErrors, string.Join("\n", refusing.Streams.Error));
                Assert.True(skipped.Skipped);
                Assert.Equal("OutputExists", skipped.SkippedReason);
                Assert.Equal(existing, File.ReadAllBytes(output));
            }
            using var replacing = PowerShell.Create(); replacing.Runspace = runspace;
            replacing.AddCommand("Convert-OfficeDocument").AddParameter("To", "Pdf").AddParameter("Path", input).AddParameter("OutputPath", output);
            replacing.Invoke(); Assert.False(replacing.HadErrors, string.Join("\n", replacing.Streams.Error));
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
            state.Commands.Add(new SessionStateCmdletEntry("Convert-OfficeDocument", typeof(ConvertOfficeDocumentCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            void AddCommand() => shell.AddCommand("Convert-OfficeDocument").AddParameter("To", "Pdf").AddParameter("Path", input)
                .AddParameter("OutputDirectory", output).AddParameter("Resume").AddParameter("CheckpointPath", Path.Combine(checkpoint, "batch.json"));
            AddCommand(); shell.AddParameter("WhatIf"); shell.Invoke();
            Assert.False(Directory.Exists(output)); Assert.False(Directory.Exists(checkpoint));
            shell.Commands.Clear(); AddCommand();
            Assert.Single(shell.Invoke()); Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error));
            shell.Commands.Clear(); AddCommand();
            var resumed = InvokeSummary(shell, runspace);
            Assert.Equal(1, resumed.Reused);
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FailedBatchRetainsSummaryAndStructuredOutcomesForBothInputShapes(bool explicitPaths) {
        string root = Path.Combine(AppContext.BaseDirectory, "pswriteoffice-failed-batch-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(root, "source");
        Directory.CreateDirectory(input);
        try {
            string[] paths = Enumerable.Range(0, 103).Select(index => Path.Combine(input, index + ".docx")).ToArray();
            foreach (string path in paths) {
                File.WriteAllText(path, "Malformed package fixture");
            }
            var state = InitialSessionState.CreateDefault();
            state.Commands.Add(new SessionStateCmdletEntry("Convert-OfficeDocument", typeof(ConvertOfficeDocumentCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            shell.AddCommand("Convert-OfficeDocument").AddParameter("To", "Pdf").AddParameter("OutputDirectory", Path.Combine(root, "output"))
                .AddParameter("ThrottleLimit", 4).AddParameter("SummaryVariable", "summary");
            if (explicitPaths) {
                shell.AddParameter("Path", paths);
            } else {
                shell.AddParameter("Path", input);
            }
            var output = shell.Invoke();
            Assert.True(shell.HadErrors);
            Assert.StartsWith("OfficeConversionJobFailed", Assert.Single(shell.Streams.Error).FullyQualifiedErrorId);
            var summary = Assert.IsType<OfficeIMO.OfficeBatchSummary>(runspace.SessionStateProxy.GetVariable("summary"));
            Assert.Equal(103, summary.Selected); Assert.Equal(103, summary.Failed); Assert.Equal(0, summary.Completed);
            var outcomes = output.Select(item => Assert.IsType<OfficeBatchItemReport>(item.BaseObject)).ToArray();
            Assert.Equal(103, outcomes.Length);
            Assert.Equal(paths.OrderBy(path => path, StringComparer.Ordinal),
                outcomes.Select(item => item.InputPath).OrderBy(path => path, StringComparer.Ordinal));
            Assert.All(outcomes, item => {
                Assert.Equal(OfficeBatchItemStatus.Failed, item.Status);
                Assert.NotEmpty(item.Diagnostics);
            });
            Assert.Empty(shell.Streams.Warning);
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
            state.Commands.Add(new SessionStateCmdletEntry("Convert-OfficeDocument", typeof(ConvertOfficeDocumentCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            var settings = new PdfTextImportOptions { PdfOptions = new PdfCreateOptions().SetEncryption("Synthetic output reader", "Synthetic output owner") };
            void AddCommand() => shell.AddCommand("Convert-OfficeDocument").AddParameter("To", "Pdf").AddParameter("Path", new[] { input })
                .AddParameter("OutputDirectory", output).AddParameter("Resume").AddParameter("CheckpointPath", Path.Combine(checkpoint, "batch.json")).AddParameter("Options", settings);
            AddCommand();
            var first = InvokeSummary(shell, runspace);
            Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error)); Assert.Equal(1, first.Completed); Assert.Equal(0, first.Failed);
            var encrypted = PdfDocument.Load(Path.Combine(output, "one.pdf"),
                new PdfLoadOptions { Password = "Synthetic output reader" });
            Assert.True(encrypted.Inspect().Security.HasEncryption); Assert.Single(encrypted.Read().Pages);
            shell.Commands.Clear(); AddCommand();
            Assert.Equal(1, InvokeSummary(shell, runspace).Reused);
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
            state.Commands.Add(new SessionStateCmdletEntry("Convert-OfficeDocument", typeof(ConvertOfficeDocumentCommand), null));
            state.Commands.Add(new SessionStateCmdletEntry("New-OfficeHtmlPdfOptions", typeof(NewOfficeHtmlPdfOptionsCommand), null));
            using var runspace = RunspaceFactory.CreateRunspace(state); runspace.Open();
            using var shell = PowerShell.Create(); shell.Runspace = runspace;
            runspace.SessionStateProxy.SetVariable("source", root);
            runspace.SessionStateProxy.SetVariable("destination", Path.Combine(root, "output"));
            runspace.SessionStateProxy.SetVariable("textSettings", new PdfTextImportOptions {
                PdfOptions = new PdfCreateOptions { PageSetup = { Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(240), global::OfficeIMO.OfficeLength.Points(180)) }, DefaultFontSize = global::OfficeIMO.OfficeLength.Points(14) }
            });
            shell.AddScript("$htmlSettings = New-OfficeHtmlPdfOptions -Margin 12 -NoInteractiveFormControls; Get-ChildItem -LiteralPath $source -File | Convert-OfficeDocument -OutputDirectory $destination -To Pdf -Options $textSettings,$htmlSettings -SummaryVariable summary");
            var results = shell.Invoke();
            Assert.False(shell.HadErrors, string.Join("\n", shell.Streams.Error));
            Assert.Equal(3, results.Count);
            Assert.Equal(2, results.Count(item => item.BaseObject is OfficeConversionReport));
            var skipped = Assert.IsType<OfficeBatchItemReport>(Assert.Single(results,
                item => item.BaseObject is OfficeBatchItemReport).BaseObject);
            Assert.True(skipped.Skipped);
            var summary = Assert.IsType<OfficeIMO.OfficeBatchSummary>(runspace.SessionStateProxy.GetVariable("summary"));
            Assert.Equal(2, summary.Completed); Assert.Equal(1, summary.Skipped); Assert.Equal(0, summary.Failed);
            var textPdf = PdfDocument.Load(Path.Combine(root, "output", "one.pdf")).Read();
            Assert.NotEmpty(textPdf.Pages);
            Assert.All(textPdf.Pages, page => Assert.Equal(240, page.Width));
            Assert.Contains("HTML pipeline text", PdfDocument.Load(Path.Combine(root, "output", "two.pdf")).Read().Text);
        } finally { Directory.Delete(root, true); }
    }
}
