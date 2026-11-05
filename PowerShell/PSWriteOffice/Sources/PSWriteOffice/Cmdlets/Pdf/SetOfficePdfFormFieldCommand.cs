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

/// <summary>Fills existing AcroForm fields from a hashtable, optionally flattening the result.</summary>
/// <para>Use Open-OfficePdf editing blocks for existing fields. Set initial field values with Add-OfficePdfFormField when composing a new PDF.</para>
/// <example><summary>Use the native OfficePdfFormField command.</summary><code>Set-OfficePdfFormField -Path .\Forms\Access-Request.pdf -OutputPath .\Forms\Access-Request-Ada.pdf -Value @{ Requester = 'Ada Lovelace'; System = 'ERP' }</code></example>
[Cmdlet("Set", "OfficePdfFormField", DefaultParameterSetName = "Document", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo), typeof(PdfDocument))]
public sealed class SetOfficePdfFormFieldCommand : OfficeCmdlet {
    private readonly List<string> _inputs = new();
    /// <summary>Path for this operation.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true, Position = 0)]
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true, Position = 0)]
    public string[]? Path { get; set; }
    /// <summary>OutputPath for this operation.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true)]
    public string? OutputPath { get; set; }
    /// <summary>Field name -> value; replaces -Field.</summary>
    [Parameter(ParameterSetName = "Path", Mandatory = true)]
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "Document", Mandatory = true)]
    public System.Collections.Hashtable? Value { get; set; }
    /// <summary>Flatten for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter Flatten { get; set; }
    /// <summary>KeepNeedAppearances for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter KeepNeedAppearances { get; set; }
    /// <summary>Incremental for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter Incremental { get; set; }
    /// <summary>AppearanceFontPath for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "Document")]
    public string? AppearanceFontPath { get; set; }
    /// <summary>AppearanceFont for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "Document")]
    public string? AppearanceFont { get; set; }
    /// <summary>Argument transformation also accepts a plain string (with a warning) or a PSCredential (its password).</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [OfficePasswordTransformation]
    public System.Security.SecureString? Password { get; set; }
    /// <summary>After successful authentication only; never bypasses a missing password.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>LoadOptions for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [OfficeOptionsTransformation(typeof(PdfLoadOptions))]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>PassThru for this operation.</summary>
    [Parameter(ParameterSetName = "Path")]
    [Parameter(ParameterSetName = "PathDirectory")]
    [Parameter(ParameterSetName = "LiteralPath")]
    [Parameter(ParameterSetName = "LiteralPathDirectory")]
    [Parameter(ParameterSetName = "Document")]
    public SwitchParameter PassThru { get; set; }
    /// <summary>OutputDirectory for this operation.</summary>
    [Parameter(ParameterSetName = "PathDirectory", Mandatory = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true)]
    public string? OutputDirectory { get; set; }
    /// <summary>LiteralPath for this operation.</summary>
    [Parameter(ParameterSetName = "LiteralPath", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Parameter(ParameterSetName = "LiteralPathDirectory", Mandatory = true, ValueFromPipelineByPropertyName = true)]
    [Alias("PSPath", "FullName")]
    public string[]? LiteralPath { get; set; }
    /// <summary>Document for this operation.</summary>
    [Parameter(ParameterSetName = "Document", ValueFromPipeline = true)]
    public PdfDocument? Document { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Incremental && (Flatten || AppearanceFontPath != null)) throw new PSArgumentException("Incremental revisions cannot flatten fields or use a custom appearance font.");
        var values = PdfCommandUtilities.ConvertFieldValues(Value!);
        if (ParameterSetName == "Document") {
            var document = Document ?? OfficeDocumentContext.Find<PdfDocument>()
                ?? throw new PSArgumentException("Provide Document or use an Open-OfficePdf content block.");
            if (!ShouldProcess("PDF document", "Fill PDF form fields")) return;
            var changed = Fill(document, values);
            bool adopted = OfficeDocumentContext.Replace(document, changed);
            if (PassThru || !adopted) WriteObject(changed);
            return;
        }
        _inputs.AddRange(OfficePath.ResolveInputs(this, LiteralPath ?? Path ?? Array.Empty<string>(), LiteralPath != null));
    }

    /// <inheritdoc />
    protected override void EndProcessing() {
        if (_inputs.Count == 0) return;
        var values = PdfCommandUtilities.ConvertFieldValues(Value!);
        foreach (var pair in PdfFileCommandInputs.Resolve(this, null, _inputs.ToArray(), OutputPath, OutputDirectory, false)) {
            if (!ShouldProcess(pair.Output, "Fill PDF form fields")) continue;
            using var document = PdfDocument.Load(pair.Input, PdfCommandUtilities.CreateReadOptions(LoadOptions, Password, IgnorePermissionRestrictions));
            using var changed = Fill(document, values);
            PdfCommandUtilities.EnsureDirectory(pair.Output);
            changed.Save(pair.Output);
            if (PassThru) WriteObject(new FileInfo(pair.Output));
        }
    }
    private PdfDocument Fill(PdfDocument document, IReadOnlyDictionary<string, string> values) {
        if (Incremental) return document.Forms.AppendRevision(values, new PdfIncrementalFormFieldUpdateOptions {
            KeepNeedAppearances = KeepNeedAppearances, GenerateAppearanceStreams = !KeepNeedAppearances
        });
        var options = PdfCommandUtilities.CreateFormFillerOptions(this, AppearanceFontPath, AppearanceFont, KeepNeedAppearances);
        if (Flatten) return options == null ? document.Forms.FillAndFlatten(values) : document.Forms.FillAndFlatten(values, options);
        return options == null ? document.Forms.Fill(values) : document.Forms.Fill(values, options);
    }
}
