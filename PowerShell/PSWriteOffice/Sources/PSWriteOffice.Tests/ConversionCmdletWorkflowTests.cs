using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PSWriteOffice.Cmdlets.Conversion;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed partial class ConversionCmdletWorkflowTests {
    [Fact]
    public void NoClobberReportsASkippedFileWithoutReplacingItsExistingDestination() {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-no-clobber-").FullName;
        try {
            string source = Path.Combine(root, "review.md"), destination = Path.Combine(root, "review.pdf");
            File.WriteAllText(source, "# Updated source");
            byte[] original = [1, 2, 3, 4];
            File.WriteAllBytes(destination, original);
            DateTime timestamp = new(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            File.SetLastWriteTimeUtc(destination, timestamp);
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                Convert-OfficeDocument -LiteralPath $source -OutputPath $destination -NoClobber -SummaryVariable result -ErrorAction Stop
                $result
                """);
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            var skipped = Assert.IsType<OfficeBatchItemReport>(output[0].BaseObject);
            Assert.True(skipped.Skipped);
            Assert.Null(skipped.Status);
            Assert.Null(skipped.OutputPath);
            Assert.Null(skipped.Report);
            Assert.Equal("OutputExists", skipped.SkippedReason);
            Assert.Same(skipped, output[1].BaseObject);
            Assert.Equal(original, File.ReadAllBytes(destination));
            Assert.Equal(timestamp, File.GetLastWriteTimeUtc(destination));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void SavedBatchKeepsItsSkipPolicyAndReportsEveryEligibleSkippedInput() {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-skip-job-").FullName;
        try {
            string source = Path.Combine(root, "review.md");
            string outputDirectory = Directory.CreateDirectory(Path.Combine(root, "pdf")).FullName;
            string destination = Path.Combine(outputDirectory, "review.pdf");
            string jobPath = Path.Combine(root, "review.officejob.json");
            string reportPath = Path.Combine(root, "run.json");
            File.WriteAllText(source, "# Saved source");
            byte[] original = [9, 8, 7, 6];
            File.WriteAllBytes(destination, original);
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("outputDirectory", outputDirectory);
            runspace.SessionStateProxy.SetVariable("jobPath", jobPath);
            runspace.SessionStateProxy.SetVariable("reportPath", reportPath);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                Convert-OfficeDocument -LiteralPath $source -To Pdf -OutputDirectory $outputDirectory -NoClobber -SkipUpToDate -SaveJob $jobPath -ErrorAction Stop | Out-Null
                Convert-OfficeDocument -JobPath $jobPath -Plan -ErrorAction Stop
                Convert-OfficeDocument -JobPath $jobPath -ReportPath $reportPath -SummaryVariable summary -ErrorAction Stop
                $summary
                """);
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            Assert.Equal(3, output.Count);
            Assert.Equal(1, Assert.IsType<OfficeBatchPlan>(output[0].BaseObject).Skipped);
            Assert.Equal("OutputExists", Assert.IsType<OfficeBatchItemReport>(output[1].BaseObject).SkippedReason);
            var summary = Assert.IsType<OfficeBatchSummary>(output[2].BaseObject);
            Assert.Equal(1, summary.Selected);
            Assert.Equal(1, summary.Skipped);
            Assert.Equal(0, summary.Completed);
            Assert.Equal(0, summary.Failed);
            var saved = OfficeBatchOptions.Load(jobPath);
            Assert.True(saved.SkipExistingOutputs);
            Assert.True(saved.SkipUpToDate);
            var report = OfficeBatchReport.Load(reportPath);
            Assert.True(report.Succeeded);
            Assert.Equal("OutputExists", Assert.Single(report.Items).SkippedReason);
            Assert.Equal(original, File.ReadAllBytes(destination));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(0)]
    public void FileFreshnessUsesStrictlyNewerOutputsAndReportsAnActualConversionOtherwise(int outputAge) {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-freshness-").FullName;
        try {
            string source = Path.Combine(root, "review.md"), destination = Path.Combine(root, "review.pdf");
            File.WriteAllText(source, "# Current conversion evidence");
            File.WriteAllBytes(destination, [1, 2, 3, 4]);
            DateTime timestamp = new(2025, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            File.SetLastWriteTimeUtc(source, timestamp);
            File.SetLastWriteTimeUtc(destination, timestamp.AddDays(outputAge));
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("Convert-OfficeDocument -LiteralPath $source -OutputPath $destination -SkipUpToDate -ErrorAction Stop");
            var item = Assert.Single(shell.Invoke()).BaseObject;
            Assert.Empty(shell.Streams.Error);
            if (outputAge > 0) Assert.Equal("UpToDate", Assert.IsType<OfficeBatchItemReport>(item).SkippedReason);
            else {
                Assert.IsType<OfficeConversionReport>(item);
                using var pdf = PdfDocument.Load(destination);
                Assert.Contains("Current conversion evidence", pdf.Read().Text);
            }
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void RegisteredChmTextAndEpubRoutesHonorTopicOptions() {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-chm-").FullName;
        try {
            string source = Path.Combine(root, "guide.chm"), destination = Path.Combine(root, "guide.epub");
            File.WriteAllBytes(source, OfficeIMO.Chm.Tests.ChmFixture.Book());
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                Convert-OfficeDocument -LiteralPath $source -To Markdown -AsString -Options @{ TopicPaths = @('/welcome.html') } -ErrorAction Stop
                Convert-OfficeDocument -LiteralPath $source -OutputPath $destination -Options @{ Chm = @{ TopicPaths = @('/welcome.html') } } -ErrorAction Stop
                """);
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            string markdown = Assert.IsType<string>(output[0].BaseObject);
            Assert.Contains("Café and help.", markdown);
            Assert.DoesNotContain("Second topic.", markdown);
            Assert.Equal(destination, Assert.IsType<OfficeConversionReport>(output[1].BaseObject).OutputPath);
            using var epub = System.IO.Compression.ZipFile.OpenRead(destination);
            string content = string.Join("\n", epub.Entries.Where(entry => entry.FullName.EndsWith(".xhtml", StringComparison.OrdinalIgnoreCase))
                .Select(entry => { using var reader = new StreamReader(entry.Open()); return reader.ReadToEnd(); }));
            Assert.Contains("Café and help.", content);
            Assert.DoesNotContain("Second topic.", content);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void SourceOverrideSelectsMarkdownInATextFileAndSurvivesPortableJobSaveAndExecution() {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-from-").FullName;
        try {
            string source = Path.Combine(root, "review.txt"), destination = Path.Combine(root, "review.docx");
            string batchOutput = Path.Combine(root, "word"), jobPath = Path.Combine(root, "review.officejob.json");
            File.WriteAllText(source, "# Forced heading\n\nSource override evidence.");
            byte[] before = File.ReadAllBytes(source);
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            runspace.SessionStateProxy.SetVariable("batchOutput", batchOutput);
            runspace.SessionStateProxy.SetVariable("jobPath", jobPath);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                Convert-OfficeDocument -LiteralPath $source -From Markdown -OutputPath $destination -ErrorAction Stop
                Get-Item -LiteralPath $source | Convert-OfficeDocument -From Markdown -To Html -AsString -ErrorAction Stop
                Convert-OfficeDocument -LiteralPath $source -From Markdown -To Word -OutputDirectory $batchOutput -SaveJob $jobPath -ErrorAction Stop | Out-Null
                Convert-OfficeDocument -JobPath $jobPath -ErrorAction Stop
                """);
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            Assert.Equal(3, output.Count);
            Assert.Equal(source, Assert.IsType<OfficeConversionReport>(output[0].BaseObject).Source);
            Assert.Contains("<h1", Assert.IsType<string>(output[1].BaseObject));
            var batchReport = Assert.IsType<OfficeConversionReport>(output[2].BaseObject);
            using var word = WordDocument.Load(batchReport.OutputPath!);
            Assert.Contains(word.Paragraphs, paragraph => paragraph.Text == "Forced heading");
            Assert.Equal(OfficeDocumentFormat.Markdown, OfficeBatchOptions.Load(jobPath).ConversionOptions.SourceFormat);
            Assert.Equal(before, File.ReadAllBytes(source));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void LiteralTextReturnsHtmlOrPublishesPdfThroughTheSameNativeCommand() {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-text-").FullName;
        try {
            string destination = Path.Combine(root, "text.pdf");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("destination", destination);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                Convert-OfficeDocument -Text '# Source text' -From Markdown -To Html -ErrorAction Stop
                Convert-OfficeDocument -Text "# Source text`n`nPublished text evidence." -From Markdown -To Pdf -OutputPath $destination -SummaryVariable result -ErrorAction Stop
                $result
                """);
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            Assert.Equal(3, output.Count);
            Assert.Contains("Source text</h1>", Assert.IsType<string>(output[0].BaseObject));
            Assert.Equal(destination, Assert.IsType<OfficeConversionReport>(output[1].BaseObject).OutputPath);
            Assert.Same(output[1].BaseObject, output[2].BaseObject);
            using var pdf = PdfDocument.Load(destination);
            Assert.Contains("Published text evidence.", pdf.Read().Text);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void FileTextExportStreamsConvertedStringsAndAssignsItsSummaryBeforeAFailingItemStopsThePipeline() {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-file-text-").FullName;
        try {
            string source = Path.Combine(root, "review[1].docx"), invalid = Path.Combine(root, "broken.docx");
            using (var document = WordDocument.Create()) { document.AddParagraph("Text pipeline evidence"); document.Save(source); }
            File.WriteAllText(invalid, "Not a document package");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("invalid", invalid);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                Get-Item -LiteralPath $source | Convert-OfficeDocument -To Markdown -AsString -SummaryVariable success -ErrorAction Stop
                $success
                $stopped = $false
                try { Convert-OfficeDocument -LiteralPath $source, $invalid -To Markdown -AsString -SummaryVariable mixed -ThrottleLimit 1 -ErrorAction Stop }
                catch { $stopped = $true }
                if (-not $stopped) { throw 'A failed input did not honor ErrorAction Stop' }
                $mixed
                """);
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            Assert.Equal(4, output.Count);
            Assert.Contains("Text pipeline evidence", Assert.IsType<string>(output[0].BaseObject));
            Assert.Equal(1, Assert.IsType<OfficeBatchSummary>(output[1].BaseObject).Completed);
            Assert.Contains("Text pipeline evidence", Assert.IsType<string>(output[2].BaseObject));
            var summary = Assert.IsType<OfficeBatchSummary>(output[3].BaseObject);
            Assert.Equal(2, summary.Selected);
            Assert.Equal(1, summary.Completed);
            Assert.Equal(1, summary.Failed);
            Assert.Single(Directory.GetFiles(root, "*.docx"), path => path == source);
            Assert.Empty(Directory.GetFiles(root, "*.md"));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void TextConversionAdmitsBytesBeforePublicationAndWhatIfCreatesNoOutput() {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-text-limits-").FullName;
        try {
            string destination = Path.Combine(root, "new", "text.pdf");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("destination", destination);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                Convert-OfficeDocument -Text 'In-memory source' -From Text -To Pdf -OutputPath $destination -WhatIf -ErrorAction Stop
                $bounded = $false
                try { Convert-OfficeDocument -Text 'Too many bytes' -From Text -To Pdf -OutputPath $destination -MaximumInputBytes 1 -ErrorAction Stop }
                catch { $bounded = $_.Exception.Message -like '*exceeds MaximumInputBytes*' }
                if (-not $bounded) { throw 'The source byte limit was not enforced' }
                if (Test-Path -LiteralPath (Split-Path $destination)) { throw 'A rejected conversion created output' }
                """);
            Assert.Empty(shell.Invoke());
            Assert.Empty(shell.Streams.Error);
            Assert.False(Directory.Exists(Path.GetDirectoryName(destination)));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void MixedFormatPaperOverrideRetainsTypedWordSettingsAndUsesEachEnginesNativePaperPreset() {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-mixed-options-").FullName;
        try {
            string input = Directory.CreateDirectory(Path.Combine(root, "inputs")).FullName;
            using (var document = WordDocument.Create()) { document.AddParagraph("Word conversion source"); document.Save(Path.Combine(input, "review.docx")); }
            File.WriteAllText(Path.Combine(input, "review.html"), "<!doctype html><html><body><p>HTML conversion source</p></body></html>");
            var settings = new WordPdfExportOptions { IncludePageNumbers = true, PageNumberFormat = "Retained page {current} of {total}" };
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("sourceDirectory", input);
            runspace.SessionStateProxy.SetVariable("destination", Path.Combine(root, "pdf"));
            runspace.SessionStateProxy.SetVariable("wordSettings", settings);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                Convert-OfficeDocument -LiteralPath $sourceDirectory -To Pdf -OutputDirectory $destination -Options $wordSettings, @{ PageSize = 'A4' } -AutoRename -ErrorAction Stop
                """);
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            Assert.Equal(2, output.Count);
            foreach (var item in output) {
                var report = Assert.IsType<OfficeConversionReport>(item.BaseObject);
                using var pdf = PdfDocument.Load(report.OutputPath!);
                var page = Assert.Single(pdf.Inspect().Pages);
                Assert.InRange(page.Width, 595, 596);
                Assert.InRange(page.Height, 841, 843);
                if (report.Source.EndsWith(".docx", StringComparison.Ordinal)) Assert.Contains("Retained page 1 of 1", pdf.Read().Text);
            }
            Assert.Null(settings.PageSize);
            Assert.Equal("Retained page {current} of {total}", settings.PageNumberFormat);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void LiteralFilePipelineConvertsToNativePdfAndRetainsOptionsAndBatchCounts() {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-command-").FullName;
        try {
            string input = Directory.CreateDirectory(Path.Combine(root, "inputs")).FullName;
            string source = Path.Combine(input, "review[1].docx");
            using (var document = WordDocument.Create()) { document.AddParagraph("Literal conversion source"); document.Save(source); }
            var options = new WordPdfExportOptions { IncludePageNumbers = true };
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("output", Path.Combine(root, "pdf"));
            runspace.SessionStateProxy.SetVariable("settings", options);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                $results = Get-Item -LiteralPath $source | Convert-OfficeDocument -To Pdf -OutputDirectory $output -Options $settings -SummaryVariable batch -ErrorAction Stop
                $results
                $batch
                """);
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            Assert.Equal(2, output.Count);
            var report = Assert.IsType<OfficeConversionReport>(output[0].BaseObject);
            var summary = Assert.IsType<OfficeBatchSummary>(output[1].BaseObject);
            Assert.Equal(1, summary.Completed);
            Assert.Equal(0, summary.Failed);
            Assert.True(options.IncludePageNumbers);
            Assert.Equal(source, report.Source);
            using var pdf = PdfDocument.Load(report.OutputPath!);
            Assert.Contains("Literal conversion source", pdf.Read().Text);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void SavedJobPlansWithoutEffectsAndRunsWithTheSharedPortableContract() {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-job-").FullName;
        try {
            string input = Directory.CreateDirectory(Path.Combine(root, "inputs")).FullName;
            using (var document = WordDocument.Create()) { document.AddParagraph("Saved conversion source"); document.Save(Path.Combine(input, "review.docx")); }
            string destination = Path.Combine(root, "pdf"), job = Path.Combine(root, "review.officejob.json");
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("inputDirectory", input);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            runspace.SessionStateProxy.SetVariable("job", job);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                $saved = Convert-OfficeDocument -LiteralPath $inputDirectory -To Pdf -OutputDirectory $destination -Options @{ PageSize = 'A4'; IncludePageNumbers = $true } -SaveJob $job -NoClobber -ErrorAction Stop
                if ($saved.FullName -ne $job) { throw ('The saved job path was lost: ' + ($Error | Out-String)) }
                $plan = Convert-OfficeDocument -JobPath $job -Plan -ErrorAction Stop
                if (-not $plan.CanRun -or $plan.Selected -ne 1) { throw 'The shared job plan was not runnable' }
                if (Test-Path -LiteralPath $destination) { throw 'Planning created the destination' }
                Convert-OfficeDocument -JobPath $job -SummaryVariable batch -ErrorAction Stop
                $batch
                """);
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            Assert.Equal(2, output.Count);
            var report = Assert.IsType<OfficeConversionReport>(output[0].BaseObject);
            Assert.Equal(1, Assert.IsType<OfficeBatchSummary>(output[1].BaseObject).Completed);
            Assert.Equal(".pdf", OfficeBatchOptions.Load(job).TargetExtension);
            using var pdf = PdfDocument.Load(report.OutputPath!);
            Assert.Contains("Saved conversion source", pdf.Read().Text);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void SingleFileConversionHonorsWhatIfAndRejectsUnusedOptionsAndDestinationMismatch() {
        string root = Directory.CreateTempSubdirectory("officeimo-conversion-policy-").FullName;
        try {
            string source = Path.Combine(root, "source.docx"), destination = Path.Combine(root, "new", "review.pdf");
            using (var document = WordDocument.Create()) { document.AddParagraph("Policy source"); document.Save(source); }
            using var runspace = OpenRunspace();
            runspace.SessionStateProxy.SetVariable("source", source);
            runspace.SessionStateProxy.SetVariable("destination", destination);
            using var shell = PowerShell.Create();
            shell.Runspace = runspace;
            shell.AddScript("""
                Convert-OfficeDocument -LiteralPath $source -OutputPath $destination -WhatIf -ErrorAction Stop
                if (Test-Path -LiteralPath (Split-Path $destination)) { throw 'WhatIf wrote output' }
                $unused = $false
                try { Convert-OfficeDocument -LiteralPath $source -OutputPath $destination -Options @{ UnknownSetting = 1 } -ErrorAction Stop }
                catch { $unused = $_.Exception.Message -like '*No selected conversion route accepts option*' }
                if (-not $unused) { throw 'Unused conversion settings were accepted' }
                $mismatch = $false
                try { Convert-OfficeDocument -LiteralPath $source -OutputPath $destination -To Word -ErrorAction Stop }
                catch { $mismatch = $_.Exception.Message -like '*same target format*' }
                if (-not $mismatch) { throw 'Mismatched target was accepted' }
                Convert-OfficeDocument -LiteralPath $source -OutputPath $destination -ErrorAction Stop
                """);
            var output = shell.Invoke();
            Assert.Empty(shell.Streams.Error);
            Assert.IsType<OfficeConversionReport>(Assert.Single(output).BaseObject);
            Assert.True(File.Exists(destination));
        } finally { Directory.Delete(root, true); }
    }

    private static Runspace OpenRunspace() {
        var state = InitialSessionState.CreateDefault();
        state.ImportPSModule(new[] { typeof(ConvertOfficeDocumentCommand).Assembly.Location });
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
}
