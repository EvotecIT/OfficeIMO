using System.Management.Automation;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates typed HTML rendering options for Export-OfficeDocumentPdf.</summary>
/// <example><summary>Export HTML with local resources and a chosen font.</summary><prefix>PS&gt; </prefix>
/// <code>$options = New-OfficeHtmlPdfOptions -FontFamily Arial -IncludeLocalResources
/// Export-OfficeDocumentPdf -InputPath .\Report.html -Path .\Report.pdf -HtmlOptions $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficeHtmlPdfOptions")]
[OutputType(typeof(HtmlToPdfOptions))]
public sealed class NewOfficeHtmlPdfOptionsCommand : PSCmdlet {
    /// <summary>Existing options to clone before applying explicitly supplied values.</summary>
    [Parameter(ValueFromPipeline = true)] public HtmlToPdfOptions? Options { get; set; }
    /// <summary>Underlying PDF writer settings, cloned before use.</summary>
    [Parameter] public PdfOptions? PdfOptions { get; set; }
    /// <summary>Default HTML font family.</summary>
    [Parameter] public string? FontFamily { get; set; }
    /// <summary>Uniform page margins in CSS pixels.</summary>
    [Parameter] [ValidateRange(0, 10000)] public double? Margin { get; set; }
    /// <summary>Allow bounded local images, stylesheets and fonts. Workflow batches constrain them to the source directory.</summary>
    [Parameter] public SwitchParameter IncludeLocalResources { get; set; }
    /// <summary>Render supported HTML form controls as interactive PDF fields.</summary>
    [Parameter] public bool? InteractiveFormControls { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = Options?.ClonePdf() ?? new HtmlToPdfOptions();
        if (PdfOptions != null) options.PdfOptions = PdfOptions.Clone();
        if (FontFamily != null) options.DefaultFontFamily = FontFamily;
        if (Margin.HasValue) options.Margins = HtmlRenderMargins.All(Margin.Value);
        if (MyInvocation.BoundParameters.ContainsKey(nameof(IncludeLocalResources))) options.ResourcePolicy.AllowLocalFileAccess = IncludeLocalResources.IsPresent;
        if (InteractiveFormControls.HasValue) options.InteractiveFormControls = InteractiveFormControls.Value;
        WriteObject(options);
    }
}
