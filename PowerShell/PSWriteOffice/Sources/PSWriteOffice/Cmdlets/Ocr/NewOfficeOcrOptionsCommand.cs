using System;
using System.Management.Automation;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Process;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Reader;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Reader;
using PSWriteOffice.Services.Reader;

namespace PSWriteOffice.Cmdlets.Ocr;

/// <summary>Creates OCR settings without starting a provider or downloading language data.</summary>
/// <para>Flat Tesseract parameters override an independent snapshot of SessionOptions. ProcessOptions and Engine select an alternative provider.</para>
/// <example><summary>Read a scanned image with selected languages.</summary>
/// <code>$ocr = New-OfficeOcrOptions -Language English,Polish -Dpi 300 -NoLanguageDownload
/// Read-OfficeDocument ./scan.png -As Text -OcrOptions $ocr</code></example>
[Cmdlet(VerbsCommon.New, "OfficeOcrOptions", DefaultParameterSetName = "Tesseract")]
[OutputType(typeof(OfficeOcrOptions))]
public sealed class NewOfficeOcrOptionsCommand : PSCmdlet {
    /// <summary>Friendly languages for the local Tesseract provider.</summary>
    [Parameter(ParameterSetName = "Tesseract"), ValidateNotNullOrEmpty]
    public TesseractOcrLanguage[]? Language { get; set; }

    /// <summary>Raw Tesseract language expression for caller-installed trained data.</summary>
    [Parameter(ParameterSetName = "Tesseract"), ValidateNotNullOrEmpty]
    public string? LanguageExpression { get; set; }

    /// <summary>Explicit local Tesseract executable path.</summary>
    [Parameter(ParameterSetName = "Tesseract"), ValidateNotNullOrEmpty]
    public string? TesseractPath { get; set; }

    /// <summary>Local directory containing installed trained-data files.</summary>
    [Parameter(ParameterSetName = "Tesseract"), ValidateNotNullOrEmpty]
    public string? DataDirectory { get; set; }

    /// <summary>Disable downloading missing curated language data.</summary>
    [Parameter(ParameterSetName = "Tesseract")]
    public SwitchParameter NoLanguageDownload { get; set; }

    /// <summary>Image resolution supplied to Tesseract.</summary>
    [Parameter(ParameterSetName = "Tesseract"), ValidateRange(1, int.MaxValue)]
    public int? Dpi { get; set; }

    /// <summary>Maximum duration of a Tesseract process, in seconds.</summary>
    [Parameter(ParameterSetName = "Tesseract"), ValidateRange(1, int.MaxValue)]
    public int? TimeoutSeconds { get; set; }

    /// <summary>Advanced Tesseract settings; flat parameters override matching values.</summary>
    [Parameter(ParameterSetName = "Tesseract"), OfficeOptionsTransformation(typeof(TesseractOcrSessionOptions)), ValidateNotNull]
    public TesseractOcrSessionOptions? SessionOptions { get; set; }

    /// <summary>Settings for the native generic JSON file-protocol OCR provider.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Process"), OfficeOptionsTransformation(typeof(ProcessOcrEngineOptions)), ValidateNotNull]
    public ProcessOcrEngineOptions? ProcessOptions { get; set; }

    /// <summary>Caller-provided OCR engine, used only when recognition is requested.</summary>
    [Parameter(Mandatory = true, ParameterSetName = "Engine"), ValidateNotNull]
    public IOcrEngine? Engine { get; set; }

    /// <summary>Bounded recognition and enrichment policy for document views.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(OfficeDocumentOcrExecutionOptions)), ValidateNotNull]
    public OfficeDocumentOcrExecutionOptions? ExecutionOptions { get; set; }

    /// <summary>Layout, confidence and resource policy for image or PDF reviews.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(PdfOcrMergeOptions)), ValidateNotNull]
    public PdfOcrMergeOptions? RecognitionOptions { get; set; }

    /// <summary>Minimum word confidence accepted into an editable OCR review.</summary>
    [Parameter, ValidateRange(0D, 1D)]
    public double? MinimumConfidence { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Language != null && LanguageExpression != null)
            throw new PSArgumentException("Use Language or LanguageExpression, not both.");
        var effective = OfficeReadOcr.Capture(new OfficeOcrOptions {
            SessionOptions = SessionOptions ?? new TesseractOcrSessionOptions(),
            ExecutionOptions = ExecutionOptions ?? new OfficeDocumentOcrExecutionOptions(),
            RecognitionOptions = RecognitionOptions ?? new PdfOcrMergeOptions { ReconstructLayout = true }
        });
        if (MinimumConfidence.HasValue) effective.RecognitionOptions.MinimumConfidence = MinimumConfidence.Value;
        if (ParameterSetName == "Process") effective.Provider = new ProcessOcrEngine(ProcessOptions!);
        else if (ParameterSetName == "Engine") effective.Provider = Engine;
        else {
            var session = effective.SessionOptions;
            if (Language != null) {
                TesseractOcrLanguage combined = 0;
                foreach (var language in Language) combined |= language;
                _ = combined.ToTesseractExpression();
                session.Languages = combined;
                session.CustomLanguageExpression = null;
                session.Engine.Language = "eng";
            }
            if (LanguageExpression != null) {
                session.CustomLanguageExpression = LanguageExpression;
                session.Engine.Language = "eng";
            }
            if (TesseractPath != null) session.Engine.ExecutablePath = OfficePath.ResolveOutput(SessionState, TesseractPath, directory: false);
            if (DataDirectory != null) session.Engine.TessdataDirectory = OfficePath.ResolveOutput(SessionState, DataDirectory, directory: true);
            if (Dpi.HasValue) session.Engine.Dpi = Dpi.Value;
            if (TimeoutSeconds.HasValue) session.Engine.Timeout = TimeSpan.FromSeconds(TimeoutSeconds.Value);
            if (MyInvocation.BoundParameters.ContainsKey(nameof(NoLanguageDownload))) session.ProvisionMissingLanguageData = !NoLanguageDownload.IsPresent;
        }
        WriteObject(effective);
    }
}
