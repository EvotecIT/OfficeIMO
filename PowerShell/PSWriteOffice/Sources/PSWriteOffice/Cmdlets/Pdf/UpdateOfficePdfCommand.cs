using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Flattens forms or annotations, or adds a searchable OCR text layer, writing a new PDF.</summary>
/// <example><summary>Use the native OfficePdf command.</summary><code>Update-OfficePdf -Path .\Forms\Access-Request-Ada.pdf -OutputPath .\Forms\Access-Request-Ada-Final.pdf -FlattenForm</code></example>
[Cmdlet("Update", "OfficePdf", DefaultParameterSetName = "Flatten", SupportsShouldProcess = true)]
[OutputType(typeof(object))]
public sealed class UpdateOfficePdfCommand : AsyncPSCmdlet {
    private readonly List<string> _inputs = new();
    /// <summary>Path for this operation.</summary>
    [Parameter(ParameterSetName = "Flatten", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "FlattenDirectory", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "Searchable", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "SearchableDirectory", Mandatory = true, Position = 0)]
    public string[]? Path { get; set; }
    /// <summary>OutputPath for this operation.</summary>
    [Parameter(ParameterSetName = "Flatten", Mandatory = true)]
    [Parameter(ParameterSetName = "FlattenLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "Searchable", Mandatory = true)]
    [Parameter(ParameterSetName = "SearchableLiteralPath", Mandatory = true)]
    public string? OutputPath { get; set; }
    /// <summary>FlattenForm for this operation.</summary>
    [Parameter(ParameterSetName = "Flatten")]
    [Parameter(ParameterSetName = "FlattenDirectory")]
    [Parameter(ParameterSetName = "FlattenLiteralPath")]
    [Parameter(ParameterSetName = "FlattenLiteralPathDirectory")]
    public SwitchParameter FlattenForm { get; set; }
    /// <summary>FlattenAnnotation for this operation.</summary>
    [Parameter(ParameterSetName = "Flatten")]
    [Parameter(ParameterSetName = "FlattenDirectory")]
    [Parameter(ParameterSetName = "FlattenLiteralPath")]
    [Parameter(ParameterSetName = "FlattenLiteralPathDirectory")]
    public SwitchParameter FlattenAnnotation { get; set; }
    /// <summary>AppearanceFontPath for this operation.</summary>
    [Parameter(ParameterSetName = "Flatten")]
    [Parameter(ParameterSetName = "FlattenDirectory")]
    [Parameter(ParameterSetName = "FlattenLiteralPath")]
    [Parameter(ParameterSetName = "FlattenLiteralPathDirectory")]
    public string? AppearanceFontPath { get; set; }
    /// <summary>AppearanceFont for this operation.</summary>
    [Parameter(ParameterSetName = "Flatten")]
    [Parameter(ParameterSetName = "FlattenDirectory")]
    [Parameter(ParameterSetName = "FlattenLiteralPath")]
    [Parameter(ParameterSetName = "FlattenLiteralPathDirectory")]
    public string? AppearanceFont { get; set; }
    /// <summary>Argument transformation also accepts a plain string (with a warning) or a PSCredential (its password).</summary>
    [Parameter(ParameterSetName = "Flatten")]
    [Parameter(ParameterSetName = "FlattenDirectory")]
    [Parameter(ParameterSetName = "FlattenLiteralPath")]
    [Parameter(ParameterSetName = "FlattenLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Searchable")]
    [Parameter(ParameterSetName = "SearchableDirectory")]
    [Parameter(ParameterSetName = "SearchableLiteralPath")]
    [Parameter(ParameterSetName = "SearchableLiteralPathDirectory")]
    [OfficePasswordTransformation]
    public System.Security.SecureString? Password { get; set; }
    /// <summary>After successful authentication only; never bypasses a missing password.</summary>
    [Parameter(ParameterSetName = "Flatten")]
    [Parameter(ParameterSetName = "FlattenDirectory")]
    [Parameter(ParameterSetName = "FlattenLiteralPath")]
    [Parameter(ParameterSetName = "FlattenLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Searchable")]
    [Parameter(ParameterSetName = "SearchableDirectory")]
    [Parameter(ParameterSetName = "SearchableLiteralPath")]
    [Parameter(ParameterSetName = "SearchableLiteralPathDirectory")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>LoadOptions for this operation.</summary>
    [Parameter(ParameterSetName = "Flatten")]
    [Parameter(ParameterSetName = "FlattenDirectory")]
    [Parameter(ParameterSetName = "FlattenLiteralPath")]
    [Parameter(ParameterSetName = "FlattenLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Searchable")]
    [Parameter(ParameterSetName = "SearchableDirectory")]
    [Parameter(ParameterSetName = "SearchableLiteralPath")]
    [Parameter(ParameterSetName = "SearchableLiteralPathDirectory")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>NoClobber for this operation.</summary>
    [Parameter(ParameterSetName = "Flatten")]
    [Parameter(ParameterSetName = "FlattenDirectory")]
    [Parameter(ParameterSetName = "FlattenLiteralPath")]
    [Parameter(ParameterSetName = "FlattenLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Searchable")]
    [Parameter(ParameterSetName = "SearchableDirectory")]
    [Parameter(ParameterSetName = "SearchableLiteralPath")]
    [Parameter(ParameterSetName = "SearchableLiteralPathDirectory")]
    public SwitchParameter NoClobber { get; set; }
    /// <summary>PassThru for this operation.</summary>
    [Parameter(ParameterSetName = "Flatten")]
    [Parameter(ParameterSetName = "FlattenDirectory")]
    [Parameter(ParameterSetName = "FlattenLiteralPath")]
    [Parameter(ParameterSetName = "FlattenLiteralPathDirectory")]
    [Parameter(ParameterSetName = "Searchable")]
    [Parameter(ParameterSetName = "SearchableDirectory")]
    [Parameter(ParameterSetName = "SearchableLiteralPath")]
    [Parameter(ParameterSetName = "SearchableLiteralPathDirectory")]
    public SwitchParameter PassThru { get; set; }
    /// <summary>OutputDirectory for this operation.</summary>
    [Parameter(ParameterSetName = "FlattenDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "FlattenLiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "SearchableDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "SearchableLiteralPathDirectory", Mandatory = true)]
    public string? OutputDirectory { get; set; }
    /// <summary>LiteralPath for this operation.</summary>
    [Parameter(ParameterSetName = "FlattenLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "FlattenLiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "SearchableLiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "SearchableLiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[]? LiteralPath { get; set; }
    /// <summary>Searchable for this operation.</summary>
    [Parameter(ParameterSetName = "Searchable", Mandatory = true)]
    [Parameter(ParameterSetName = "SearchableDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "SearchableLiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "SearchableLiteralPathDirectory", Mandatory = true)]
    public SwitchParameter Searchable { get; set; }
    /// <summary>New-OfficeOcrOptions: engine, languages, Tesseract path, data directory, DPI, confidence; same type Read-/Convert-OfficeDocument -Ocr use.</summary>
    [Parameter(ParameterSetName = "Searchable")]
    [Parameter(ParameterSetName = "SearchableDirectory")]
    [Parameter(ParameterSetName = "SearchableLiteralPath")]
    [Parameter(ParameterSetName = "SearchableLiteralPathDirectory")]
    [OfficeOptionsTransformation(typeof(PSWriteOffice.Models.Reader.OfficeOcrOptions))]
    public PSWriteOffice.Models.Reader.OfficeOcrOptions? OcrOptions { get; set; }
    /// <summary>Replaces -PdfOptions.</summary>
    [Parameter(ParameterSetName = "Searchable")]
    [Parameter(ParameterSetName = "SearchableDirectory")]
    [Parameter(ParameterSetName = "SearchableLiteralPath")]
    [Parameter(ParameterSetName = "SearchableLiteralPathDirectory")]
    [OfficeOptionsTransformation(typeof(PdfOcrMergeOptions))]
    public PdfOcrMergeOptions? OcrMergeOptions { get; set; }

    /// <inheritdoc />
    protected override System.Threading.Tasks.Task ProcessRecordAsync() {
        _inputs.AddRange(OfficePath.ResolveInputs(this, LiteralPath ?? Path ?? Array.Empty<string>(), LiteralPath != null));
        return System.Threading.Tasks.Task.CompletedTask;
    }

    /// <inheritdoc />
    protected override async System.Threading.Tasks.Task EndProcessingAsync() {
        if (!Searchable && !FlattenForm && !FlattenAnnotation) throw new PSArgumentException("Select FlattenForm, FlattenAnnotation or Searchable.");
        var inputs = PdfFileCommandInputs.Resolve(this, null, _inputs.ToArray(), OutputPath, OutputDirectory, NoClobber);
        var readOptions = PdfCommandUtilities.CreateReadOptions(LoadOptions, Password, IgnorePermissionRestrictions);
        var formOptions = PdfCommandUtilities.CreateFormFillerOptions(this, AppearanceFontPath, AppearanceFont, false);
        foreach (var pair in inputs) {
            if (!ShouldProcess(pair.Output, Searchable ? "Create searchable PDF" : "Flatten PDF content")) continue;
            var document = PdfDocument.Load(pair.Input, readOptions);
            object result = document;
            if (Searchable) {
                var engine = OcrOptions?.Provider;
                if (engine == null) engine = (await OfficeIMO.Ocr.Tesseract.TesseractOcr.CreateSessionAsync(
                    OcrOptions?.SessionOptions ?? new OfficeIMO.Ocr.Tesseract.TesseractOcrSessionOptions(), CancelToken).ConfigureAwait(false)).Engine;
                var recognized = await document.MakeSearchableAsync(engine, OcrMergeOptions ?? OcrOptions?.RecognitionOptions, CancelToken).ConfigureAwait(false);
                document = recognized.Document;
                result = recognized;
            } else {
                if (FlattenForm) document = formOptions == null ? document.Forms.Flatten() : document.Forms.Flatten(formOptions);
                if (FlattenAnnotation) document = document.FlattenVisualAnnotations();
                result = document;
            }
            PdfCommandUtilities.EnsureDirectory(pair.Output);
            await document.SaveAsync(pair.Output, new PdfSaveOptions { ConflictPolicy = NoClobber ? OfficeConversionFileConflictPolicy.FailIfExists : OfficeConversionFileConflictPolicy.Replace }, CancelToken).ConfigureAwait(false);
            WriteObject(PassThru ? result : new FileInfo(pair.Output));
        }
    }
}
