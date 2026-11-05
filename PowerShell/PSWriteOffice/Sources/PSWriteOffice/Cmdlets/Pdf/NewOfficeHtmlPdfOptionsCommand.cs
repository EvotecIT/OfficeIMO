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
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates HTML-to-PDF options for Convert-OfficeDocument -To Pdf.</summary>
/// <example><summary>Use the native OfficeHtmlPdfOptions command.</summary><code>Convert-OfficeDocument -Path .\Report.html -To Pdf -OutputPath .\Report.pdf -Options (New-OfficeHtmlPdfOptions -Margin 2cm)</code></example>
[Cmdlet("New", "OfficeHtmlPdfOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(HtmlPdfExportOptions))]
public sealed class NewOfficeHtmlPdfOptionsCommand : OfficeCmdlet {
    /// <summary>BaseOptions for this operation.</summary>
    [Parameter(ParameterSetName = "Default", ValueFromPipeline = true)]
    [OfficeOptionsTransformation(typeof(HtmlPdfExportOptions))]
    public HtmlPdfExportOptions? BaseOptions { get; set; }
    /// <summary>PdfOptions for this operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeOptionsTransformation(typeof(PdfCreateOptions))]
    public PdfCreateOptions? PdfOptions { get; set; }
    /// <summary>Font for this operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Font { get; set; }
    /// <summary>Margin for this operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    [OfficeLengthTransformation]
    public OfficeLength? Margin { get; set; }
    /// <summary>IncludeLocalResources for this operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter IncludeLocalResources { get; set; }
    /// <summary>NoInteractiveFormControls for this operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public SwitchParameter NoInteractiveFormControls { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = BaseOptions?.ClonePdf() ?? new HtmlPdfExportOptions();
        if (PdfOptions != null) options.PdfOptions = PdfOptions.Clone();
        if (Font != null) options.DefaultFontFamily = Font;
        if (Margin.HasValue) options.Margins = HtmlRenderMargins.All(Margin.Value.ToPixels());
        if (BoundSwitch(nameof(IncludeLocalResources)) is bool local) options.ResourcePolicy.AllowLocalFileAccess = local;
        if (BoundSwitch(nameof(NoInteractiveFormControls)) is bool disabled) options.InteractiveFormControls = !disabled;
        WriteObject(options);
    }
}
