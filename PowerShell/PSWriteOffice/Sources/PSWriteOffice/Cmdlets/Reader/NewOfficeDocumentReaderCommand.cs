using System;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Threading.Tasks;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Process;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Reader;

/// <summary>Creates an immutable fully configured OfficeIMO document reader.</summary>
/// <example>
///   <summary>Create a reader with OCR and a resilient processor policy.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$reader = New-OfficeDocumentReader -OcrLanguage English, Polish -MaxStoreItems 5000 -ProcessorFailureBehavior ContinueWithDiagnostic</code>
///   <para>The returned reader can be supplied to every PSWriteOffice Reader command.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeDocumentReader")]
[OutputType(typeof(OfficeDocumentReader))]
public sealed class NewOfficeDocumentReaderCommand : AsyncPSCmdlet
{
    /// <summary>Advanced format-specific settings supplied by a .NET host.</summary>
    [Parameter(DontShow = true)]
    public ReaderAllOptions? ReaderAllOptions { get; set; }

    /// <summary>Additional ordered processors to run after document extraction.</summary>
    [Parameter(DontShow = true)]
    public IOfficeDocumentProcessor[]? Processor { get; set; }

    /// <summary>Caller-provided OCR engine.</summary>
    [Parameter(DontShow = true)]
    public IOcrEngine? OcrEngine { get; set; }

    /// <summary>Configure the built-in Tesseract command-line OCR adapter.</summary>
    [Parameter(DontShow = true)]
    public TesseractOcrSessionOptions? TesseractOptions { get; set; }

    /// <summary>Configure the generic JSON file-protocol OCR process adapter.</summary>
    [Parameter(DontShow = true)]
    public ProcessOcrEngineOptions? ProcessOcrOptions { get; set; }

    /// <summary>Optional OCR execution limits and merge behavior.</summary>
    [Parameter(DontShow = true)]
    public OfficeDocumentOcrExecutionOptions? OcrOptions { get; set; }

    /// <summary>Enable the built-in Tesseract command-line OCR adapter with default settings.</summary>
    [Parameter]
    public SwitchParameter UseTesseract { get; set; }

    /// <summary>Tesseract executable path or command name.</summary>
    [Parameter]
    public string? TesseractExecutablePath { get; set; }

    /// <summary>Friendly OCR languages used by the built-in Tesseract adapter.</summary>
    [Parameter]
    [ValidateNotNullOrEmpty]
    public TesseractOcrLanguage[]? OcrLanguage { get; set; }

    /// <summary>Advanced raw Tesseract expression for caller-installed custom trained-data models.</summary>
    [Parameter]
    public string? TesseractLanguage { get; set; }

    /// <summary>Optional Tesseract tessdata directory.</summary>
    [Parameter]
    public string? TesseractDataPath { get; set; }

    /// <summary>Optional input DPI passed to Tesseract.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? TesseractDpi { get; set; }

    /// <summary>Maximum Tesseract process duration in seconds. The default is 120.</summary>
    [Parameter]
    [ValidateRange(1, 86400)]
    public int? TesseractTimeoutSeconds { get; set; }

    /// <summary>Maximum PST, OST, OLM, or EMLX items projected from each store. The default is 1,000.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaxStoreItems { get; set; }

    /// <summary>Project every matching item from each email store.</summary>
    [Parameter]
    public SwitchParameter AllStoreItems { get; set; }

    /// <summary>Maximum asynchronous reads allowed in flight.</summary>
    [Parameter]
    [ValidateRange(1, 64)]
    public int? MaxConcurrentReads { get; set; }

    /// <summary>Behavior when a processor fails.</summary>
    [Parameter]
    public OfficeDocumentProcessorFailureBehavior ProcessorFailureBehavior { get; set; } = OfficeDocumentProcessorFailureBehavior.Throw;

    /// <inheritdoc />
    protected override async Task ProcessRecordAsync()
    {
        ValidatePowerShellParameters();
        var configuredOcr = new List<IOcrEngine>();
        if (OcrEngine != null) configuredOcr.Add(OcrEngine);
        if (TesseractOptions != null)
        {
            configuredOcr.Add((await TesseractOcr.CreateSessionAsync(TesseractOptions, CancelToken).ConfigureAwait(false)).Engine);
        }
        if (HasPowerShellTesseractConfiguration())
        {
            var sessionOptions = new TesseractOcrSessionOptions();
            var engineOptions = sessionOptions.Engine;
            if (!string.IsNullOrWhiteSpace(TesseractExecutablePath)) engineOptions.ExecutablePath = TesseractExecutablePath!;
            if (OcrLanguage is { Length: > 0 }) sessionOptions.Languages = CombineLanguages(OcrLanguage);
            else if (!string.IsNullOrWhiteSpace(TesseractLanguage)) sessionOptions.CustomLanguageExpression = TesseractLanguage!;
            if (!string.IsNullOrWhiteSpace(TesseractDataPath)) engineOptions.TessdataDirectory = TesseractDataPath!;
            if (TesseractDpi.HasValue) engineOptions.Dpi = TesseractDpi.Value;
            if (TesseractTimeoutSeconds.HasValue) engineOptions.Timeout = TimeSpan.FromSeconds(TesseractTimeoutSeconds.Value);
            configuredOcr.Add((await TesseractOcr.CreateSessionAsync(sessionOptions, CancelToken).ConfigureAwait(false)).Engine);
        }
        if (ProcessOcrOptions != null) configuredOcr.Add(new ProcessOcrEngine(ProcessOcrOptions));
        int? storeItemLimit = AllStoreItems.IsPresent ? int.MaxValue : MaxStoreItems;
        var powerShellConfiguration = ReaderCommandUtilities.BuildSearchConfiguration(
            includePageLocations: false,
            storeItemLimit);
        var builder = ReaderCommandUtilities.CreateBuilder(
                ReaderAllOptions ?? powerShellConfiguration.HandlerOptions)
            .WithProcessorFailureBehavior(ProcessorFailureBehavior);
        if (MaxConcurrentReads.HasValue) builder.WithMaxConcurrentReads(MaxConcurrentReads.Value);
        if (Processor != null) builder.AddProcessors(Processor.Where(static value => value != null));
        if (configuredOcr.Count == 1) builder.AddProcessor(new OfficeDocumentOcrProcessor(configuredOcr[0], OcrOptions));
        WriteObject(builder.Build());
    }

    private bool HasPowerShellTesseractConfiguration()
    {
        return UseTesseract.IsPresent ||
               !string.IsNullOrWhiteSpace(TesseractExecutablePath) ||
               OcrLanguage is { Length: > 0 } ||
               !string.IsNullOrWhiteSpace(TesseractLanguage) ||
               !string.IsNullOrWhiteSpace(TesseractDataPath) ||
               TesseractDpi.HasValue ||
               TesseractTimeoutSeconds.HasValue;
    }

    private void ValidatePowerShellParameters()
    {
        if (AllStoreItems.IsPresent && MaxStoreItems.HasValue)
        {
            throw new PSArgumentException("Specify either -MaxStoreItems or -AllStoreItems, not both.");
        }
        if (ReaderAllOptions != null && (AllStoreItems.IsPresent || MaxStoreItems.HasValue))
        {
            throw new PSArgumentException(
                "Use -MaxStoreItems/-AllStoreItems or the advanced -ReaderAllOptions object, not both.");
        }
        if (TesseractOptions != null && HasPowerShellTesseractConfiguration())
        {
            throw new PSArgumentException(
                "Use PowerShell Tesseract parameters or the advanced -TesseractOptions object, not both.");
        }
        if (OcrLanguage is { Length: > 0 } && !string.IsNullOrWhiteSpace(TesseractLanguage))
        {
            throw new PSArgumentException(
                "Use -OcrLanguage or the advanced -TesseractLanguage parameter, not both.");
        }

        var ocrConfigurationCount = 0;
        if (OcrEngine != null) ocrConfigurationCount++;
        if (TesseractOptions != null) ocrConfigurationCount++;
        if (HasPowerShellTesseractConfiguration()) ocrConfigurationCount++;
        if (ProcessOcrOptions != null) ocrConfigurationCount++;
        if (ocrConfigurationCount > 1)
        {
            throw new PSArgumentException(
                "Specify only one OCR source: -OcrEngine, -TesseractOptions, -ProcessOcrOptions, or the public Tesseract parameters.");
        }
    }

    private static TesseractOcrLanguage CombineLanguages(IEnumerable<TesseractOcrLanguage> languages)
    {
        TesseractOcrLanguage combined = 0;
        foreach (TesseractOcrLanguage language in languages) combined |= language;
        _ = combined.ToTesseractExpression();
        return combined;
    }
}
