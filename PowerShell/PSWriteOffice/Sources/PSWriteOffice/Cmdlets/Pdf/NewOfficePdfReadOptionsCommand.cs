using System.Collections;
using System.Management.Automation;
using System.Net;
using System.Security;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Models.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates a detached native PDF read policy and optional OCR-review settings.</summary>
/// <example><summary>Read selected pages with bounded untrusted-input parsing.</summary>
/// <code>$options = New-OfficePdfReadOptions -LoadOptions ([OfficeIMO.Pdf.PdfLoadOptions]::UntrustedDefaults) -Profile Fast -Page '2-4'
/// Read-OfficeDocument -Path ./Report.pdf -Options $options -As Text</code></example>
/// <example><summary>Limit rasterization before an explicitly requested OCR review.</summary>
/// <code>$ocr = New-OfficeOcrOptions -NoLanguageDownload
/// $options = New-OfficePdfReadOptions -OcrMergeOptions @{ MaxPixelsPerPage = 2000000; Dpi = 96 }
/// Read-OfficeDocument -Path ./Scan.pdf -Options $options -As Ocr -OcrOptions $ocr</code></example>
[Cmdlet(VerbsCommon.New, "OfficePdfReadOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(PdfReadOptions))]
public sealed class NewOfficePdfReadOptionsCommand : OfficeCmdlet {
    /// <summary>Semantic reconstruction profile; omission preserves the base policy.</summary>
    [Parameter(ParameterSetName = "Default")]
    public PdfReadProfile? Profile { get; set; }
    /// <summary>Caller-ordered one-based source pages.</summary>
    [Parameter, OfficePageRangeTransformation, ValidateNotNull]
    public OfficePageRange? Page { get; set; }
    /// <summary>Native layout and reading-order settings.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(PdfTextLayoutOptions)), ValidateNotNull]
    public PdfTextLayoutOptions? LayoutOptions { get; set; }
    /// <summary>Password used to open encrypted PDFs; omission preserves the base policy.</summary>
    [Parameter, OfficePasswordTransformation]
    public SecureString? Password { get; set; }
    /// <summary>Ignore permission restrictions after successful decryption; explicit false restores enforcement.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }
    /// <summary>Complete base policy, including strict parsing and resource ceilings.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(PdfLoadOptions)), ValidateNotNull]
    public PdfLoadOptions? LoadOptions { get; set; }
    /// <summary>Native OCR-review settings, used by Read-OfficeDocument -As Ocr.</summary>
    [Parameter, OfficeOptionsTransformation(typeof(PdfOcrMergeOptions)), ValidateNotNull]
    public PdfOcrMergeOptions? OcrMergeOptions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        if (Profile.HasValue && !System.Enum.IsDefined(typeof(PdfReadProfile), Profile.Value))
            throw new PSArgumentException("Unsupported PDF read profile.");
        var ocr = OcrMergeOptions?.Clone();
        var policy = (LoadOptions ?? ocr?.ReadOptions ?? new PdfLoadOptions()).Clone();
        var overrides = new Hashtable();
        if (Profile.HasValue) overrides[nameof(PdfLoadOptions.Profile)] = Profile.Value;
        if (Page != null) overrides[nameof(PdfLoadOptions.PageSelection)] = Page.Selection;
        if (LayoutOptions != null) overrides[nameof(PdfLoadOptions.LayoutOptions)] = LayoutOptions;
        if (MyInvocation.BoundParameters.ContainsKey(nameof(Password)))
            overrides[nameof(PdfLoadOptions.Password)] = Password == null ? null : new NetworkCredential(string.Empty, Password).Password;
        if (BoundSwitch(nameof(IgnorePermissionRestrictions)) is bool ignore)
            overrides[nameof(PdfLoadOptions.PermissionPolicy)] = ignore ? PdfPermissionPolicy.IgnoreRestrictions : PdfPermissionPolicy.Enforce;
        var engine = (EngineIntrinsics)SessionState.PSVariable.GetValue("ExecutionContext");
        new OfficeOptionsTransformationAttribute(typeof(PdfLoadOptions)).ApplyTo(engine, policy, overrides);
        policy = policy.Clone();
        if (ocr != null) ocr.ReadOptions = policy.Clone();
        WriteObject(new PdfReadOptions { LoadOptions = policy, OcrMergeOptions = ocr });
    }
}
