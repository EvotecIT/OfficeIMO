using System.Security.Cryptography;
using OfficeIMO.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeConverterBatchPersistenceTests {
    [Fact]
    public void PortableOptionsRemainSerializableAfterCloning() {
        using var files = new SavedFiles();
        var original = new OfficeConversionOptions {
            WordLoad = new WordLoadOptions { FailOnLoss = false },
            PlainText = new PdfTextImportOptions { TabSize = 4 }
        };
        OfficeConversionOptions cloned = original.Clone();
        var batch = files.Settings() with { ConversionOptions = cloned };
        string json = batch.ToJson(files.Root);
        OfficeBatchOptions loaded = OfficeBatchOptions.Parse(json, files.Root);
        Assert.Equal(4, loaded.ConversionOptions.PlainText!.TabSize);
        Assert.Null(cloned.Reader);
        cloned.PlainText!.TabSize = 6;
        Assert.Equal(4, original.PlainText.TabSize);
    }

    [Fact]
    public async Task SavedDefinitionResolvesAgainstItsOwnDirectoryAndRecordsExactBytes() {
        using var files = new SavedFiles();
        string jobPath = Path.Combine(files.Root, "definition.json"), reportPath = Path.Combine(files.Root, "run.json");
        OfficeBatchOptions options = files.Settings();
        options.Save(jobPath);
        byte[] json = File.ReadAllBytes(jobPath);
        byte[] bom = Encoding.UTF8.GetPreamble();
        byte[] admitted = bom.Concat(json).ToArray();
        File.WriteAllBytes(jobPath, admitted);
        OfficeBatchOptions loaded = OfficeBatchOptions.Load(jobPath);
        Assert.Equal(options.InputPaths, loaded.InputPaths);
        Assert.Equal(options.OutputDirectory, loaded.OutputDirectory);
        OfficeBatchReport result = await OfficeConverter.ConvertAsync(loaded with { ReportPath = reportPath });
        Assert.True(result.Succeeded);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(admitted)), result.JobSha256);
        OfficeBatchReport saved = OfficeBatchReport.Load(reportPath);
        Assert.Equal(OfficeBatchRunState.Completed, saved.State);
        Assert.Equal(result.Summary, saved.Summary);
        Assert.False(OfficeBatchReport.IsActive(reportPath));
        Assert.False(File.Exists(reportPath + ".lease"));
        Assert.True(OfficeBatchReport.TryDeleteFinished(reportPath));
        Assert.False(File.Exists(reportPath));
    }

    [Fact]
    public void SavedDefinitionsRejectUnknownFieldsDuplicatesAndNumericEnums() {
        using var files = new SavedFiles();
        string json = files.Settings().ToJson(files.Root);
        Assert.ThrowsAny<System.Text.Json.JsonException>(() => OfficeBatchOptions.Parse(json.Replace("\"schema\":", "\"unknown\": 1, \"schema\":"), files.Root));
        Assert.ThrowsAny<System.Text.Json.JsonException>(() => OfficeBatchOptions.Parse(json.Replace("\"schema\":", "\"schema\": \"duplicate\", \"schema\":"), files.Root));
        Assert.ThrowsAny<System.Text.Json.JsonException>(() => OfficeBatchOptions.Parse(json.Replace("\"Fail\"", "0"), files.Root));
    }

    [Fact]
    public async Task PasswordValuesNeverEnterSavedDefinitionsOrRecordedReports() {
        using var files = new SavedFiles();
        string input = Path.Combine(files.Root, "source.docx");
        const string password = "synthetic-secret-for-batch";
        using (var document = WordDocument.Create()) { document.AddParagraph("Encrypted input"); document.Save(input, new WordSaveOptions { Password = password }); }
        OfficeBatchOptions options = files.Settings() with { InputPaths = [input], SourcePasswordEnvironmentVariable = "OFFICE_TEST_PASSWORD" };
        Assert.Throws<NotSupportedException>(() => (options with { ConversionOptions = new OfficeConversionOptions { SourcePassword = password } }).ToJson(files.Root));
        string job = Path.Combine(files.Root, "definition.json"); options.Save(job);
        Assert.DoesNotContain(password, File.ReadAllText(job));
        OfficeBatchOptions loaded = OfficeBatchOptions.Load(job);
        int resolutions = 0;
        loaded = loaded with { ResolveEnvironmentVariable = name => { resolutions++; Assert.Equal("OFFICE_TEST_PASSWORD", name); return password; },
            ReportPath = Path.Combine(files.Root, "run.json") };
        Assert.True((await OfficeConverter.PlanAsync(loaded)).CanRun);
        Assert.Equal(0, resolutions);
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(loaded);
        Assert.True(report.Succeeded, string.Join("; ", report.Items.Select(item => item.Summary)));
        Assert.Equal(1, resolutions);
        Assert.DoesNotContain(password, File.ReadAllText(loaded.ReportPath));
        using var pdf = PdfDocument.Load(Assert.Single(report.Items).OutputPath!);
        Assert.Contains("Encrypted input", pdf.ExtractText());
    }

    [Fact]
    public async Task SavedPdfPasswordReferenceAndCompressionSurvivePortableRoundTrip() {
        using var files = new SavedFiles();
        const string password = "synthetic-pdf-source-password";
        string input = Path.Combine(files.Root, "encrypted.pdf");
        using (var document = PdfDocument.Create(new PdfCreateOptions().SetEncryption(new PdfStandardEncryptionOptions(password)))) {
            document.Compose(builder => builder.Content(content => content.Text("Separate PDF password reference")));
            document.Save(input);
        }
        OfficeBatchOptions options = files.Settings() with {
            InputPaths = [input], TargetExtension = ".docx", PdfPasswordEnvironmentVariable = "PDF_TEST_PASSWORD",
            ConversionOptions = new OfficeConversionOptions { FailOnLoss = false }
        };
        string job = Path.Combine(files.Root, "pdf-definition.json");
        options.Save(job);
        Assert.DoesNotContain(password, File.ReadAllText(job));
        var loaded = OfficeBatchOptions.Load(job) with { ResolveEnvironmentVariable = name => name == "PDF_TEST_PASSWORD" ? password : null };
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(loaded);
        Assert.True(report.Succeeded, string.Join("; ", report.Items.Select(item => item.Summary)));
        using var word = WordDocument.Load(Assert.Single(report.Items).OutputPath!);
        Assert.Contains("Separate PDF password reference", string.Join(" ", word.Paragraphs.Select(paragraph => paragraph.Text)));
        OfficeBatchOptions compressed = files.Settings() with { ConversionOptions = new OfficeConversionOptions { CompressPdfOutput = true } };
        compressed.Save(Path.Combine(files.Root, "compressed-definition.json"));
        Assert.True(OfficeBatchOptions.Load(Path.Combine(files.Root, "compressed-definition.json")).ConversionOptions.CompressPdfOutput);
    }

    [Fact]
    public async Task ExistingReportAndChangedSavedDefinitionRemainUntouched() {
        using var files = new SavedFiles();
        string job = Path.Combine(files.Root, "definition.json"), report = Path.Combine(files.Root, "run.json");
        files.Settings().Save(job);
        OfficeBatchOptions loaded = OfficeBatchOptions.Load(job) with { ReportPath = report };
        File.WriteAllText(job, File.ReadAllText(job) + " ");
        await Assert.ThrowsAsync<InvalidDataException>(() => OfficeConverter.ConvertAsync(loaded));
        Assert.False(Directory.Exists(loaded.OutputDirectory));
        Assert.False(File.Exists(report));
        File.WriteAllText(report, "Keep this report");
        await Assert.ThrowsAsync<IOException>(() => OfficeConverter.ConvertAsync(files.Settings() with { ReportPath = report }));
        Assert.Equal("Keep this report", File.ReadAllText(report));
        Assert.False(Directory.Exists(loaded.OutputDirectory));
    }

    [Fact]
    public async Task RetentionCanStopRecordingWithoutCancellingConversion() {
        using var files = new SavedFiles();
        string reportPath = Path.Combine(files.Root, "run.json");
        bool retain = true;
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(files.Settings() with { ReportPath = reportPath, RetainReport = () => retain },
            new CallbackProgress(_ => retain = false));
        Assert.True(report.Succeeded);
        Assert.False(File.Exists(reportPath));
        Assert.False(File.Exists(reportPath + ".lease"));
    }

    private sealed class CallbackProgress(Action<OfficeBatchItemReport> callback) : IProgress<OfficeBatchItemReport> {
        public void Report(OfficeBatchItemReport item) => callback(item);
    }
    private sealed class SavedFiles : IDisposable {
        internal string Root { get; } = Directory.CreateTempSubdirectory("officeimo-saved-batch-test-").FullName;
        internal OfficeBatchOptions Settings() {
            string input = Path.Combine(Root, "source.txt"); File.WriteAllText(input, "Saved conversion text");
            return new() { InputPaths = [input], OutputDirectory = Path.Combine(Root, "output") };
        }
        public void Dispose() => Directory.Delete(Root, true);
    }
}
