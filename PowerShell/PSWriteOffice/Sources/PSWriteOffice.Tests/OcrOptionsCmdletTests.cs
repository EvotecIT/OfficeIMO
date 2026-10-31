using System.Management.Automation;
using System.Management.Automation.Runspaces;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Process;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Cmdlets.Ocr;
using PSWriteOffice.Models.Reader;

namespace PSWriteOffice.Tests;

[Collection(PowerShellRunspaceCollection.Name)]
public sealed class OcrOptionsCmdletTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-ocr-options-" + Guid.NewGuid().ToString("N"));
    public OcrOptionsCmdletTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    [Fact]
    public void TesseractConvenienceSettingsOverrideIndependentNativeOptionsWithoutRunningOcr() {
        var source = new TesseractOcrSessionOptions { CustomLanguageExpression = "custom", ProvisionMissingLanguageData = true };
        source.Engine.Dpi = 123;
        source.Engine.Timeout = TimeSpan.FromSeconds(45);
        var recognition = new PdfOcrMergeOptions { MinimumConfidence = 0.2 };
        using var runspace = Open();
        runspace.SessionStateProxy.Path.SetLocation(_root);
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("New-OfficeOcrOptions").AddParameter("SessionOptions", source)
            .AddParameter("RecognitionOptions", recognition).AddParameter("MinimumConfidence", 0.8)
            .AddParameter("Language", new[] { TesseractOcrLanguage.English, TesseractOcrLanguage.Polish })
            .AddParameter("TesseractPath", "tools/tesseract").AddParameter("DataDirectory", "trained-data")
            .AddParameter("Dpi", 300).AddParameter("TimeoutSeconds", 12).AddParameter("NoLanguageDownload");
        var result = Assert.IsType<OfficeOcrOptions>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.Null(result.Provider);
        Assert.Equal(TesseractOcrLanguage.English | TesseractOcrLanguage.Polish, result.SessionOptions.Languages);
        Assert.Null(result.SessionOptions.CustomLanguageExpression);
        Assert.Equal(Path.Combine(_root, "tools", "tesseract"), result.SessionOptions.Engine.ExecutablePath);
        Assert.Equal(Path.Combine(_root, "trained-data"), result.SessionOptions.Engine.TessdataDirectory);
        Assert.Equal(300, result.SessionOptions.Engine.Dpi);
        Assert.Equal(TimeSpan.FromSeconds(12), result.SessionOptions.Engine.Timeout);
        Assert.False(result.SessionOptions.ProvisionMissingLanguageData);
        Assert.Equal(0.8, result.RecognitionOptions.MinimumConfidence);
        Assert.NotSame(source, result.SessionOptions);
        Assert.NotSame(source.Engine, result.SessionOptions.Engine);
        Assert.NotSame(source.LanguageData, result.SessionOptions.LanguageData);
        Assert.Equal(123, source.Engine.Dpi);
        Assert.Equal(TimeSpan.FromSeconds(45), source.Engine.Timeout);
        Assert.Equal("custom", source.CustomLanguageExpression);
        Assert.True(source.ProvisionMissingLanguageData);
        Assert.Equal(0.2, recognition.MinimumConfidence);
        Assert.Empty(Directory.EnumerateFileSystemEntries(_root));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AlternativeProviderSelectionIsInertAndRejectsTesseractOnlySettings(bool suppliedEngine) {
        var source = new ProcessOcrEngineOptions { FileName = Path.Combine(_root, "missing-ocr-provider"), Id = "configured-inert" };
        IOcrEngine engine = new ProcessOcrEngine(source);
        string parameter = suppliedEngine ? "Engine" : "ProcessOptions";
        object input = suppliedEngine ? engine : source;
        using var runspace = Open();
        using var shell = PowerShell.Create();
        shell.Runspace = runspace;
        shell.AddCommand("New-OfficeOcrOptions").AddParameter(parameter, input);
        var result = Assert.IsType<OfficeOcrOptions>(Assert.Single(shell.Invoke()).BaseObject);
        Assert.False(shell.HadErrors, Errors(shell));
        Assert.IsType<ProcessOcrEngine>(result.Provider);
        Assert.Equal("configured-inert", result.Provider!.Id);
        if (suppliedEngine) Assert.Same(engine, result.Provider);
        shell.Commands.Clear();
        shell.AddCommand("New-OfficeOcrOptions").AddParameter(parameter, input).AddParameter("Language", "English");
        Exception? failure = Record.Exception(() => shell.Invoke());
        Assert.True(failure != null || shell.HadErrors);
        Assert.Empty(Directory.EnumerateFileSystemEntries(_root));
    }

    private static Runspace Open() {
        var state = InitialSessionState.CreateDefault();
        state.Commands.Add(new SessionStateCmdletEntry("New-OfficeOcrOptions", typeof(NewOfficeOcrOptionsCommand), null));
        var runspace = RunspaceFactory.CreateRunspace(state);
        runspace.Open();
        return runspace;
    }
    private static string Errors(PowerShell shell) => string.Join(Environment.NewLine, shell.Streams.Error);
}
