using System.Management.Automation;
using OfficeIMO.Rtf.Pdf;

namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Creates discoverable RTF-to-PDF conversion options for Export-OfficeDocumentPdf.</summary>
/// <example>
///   <summary>Include document structure and bound system-font discovery.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeRtfPdfOptions -IncludeImages -IncludeTables -IncludeHeaderFooters -MaximumSystemFontFamilies 32
/// Export-OfficeDocumentPdf -InputPath .\Report.rtf -Path .\Report.pdf -RtfOptions $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeRtfPdfOptions")]
[OutputType(typeof(RtfToPdfOptions))]
public sealed class NewOfficeRtfPdfOptionsCommand : PSCmdlet {
    /// <summary>Underlying low-level OfficeIMO PDF options.</summary>
    [Parameter]
    public OfficeIMO.Pdf.PdfOptions? PdfOptions { get; set; }

    /// <summary>Include text marked hidden.</summary>
    [Parameter]
    public SwitchParameter IncludeHiddenText { get; set; }

    /// <summary>Render images.</summary>
    [Parameter]
    public SwitchParameter IncludeImages { get; set; }

    /// <summary>Fallback image width in PDF points.</summary>
    [Parameter]
    [ValidateRange(double.Epsilon, double.MaxValue)]
    public double? DefaultImageWidth { get; set; }

    /// <summary>Fallback image height in PDF points.</summary>
    [Parameter]
    [ValidateRange(double.Epsilon, double.MaxValue)]
    public double? DefaultImageHeight { get; set; }

    /// <summary>Copy document metadata into the PDF.</summary>
    [Parameter]
    public SwitchParameter IncludeMetadata { get; set; }

    /// <summary>Render tables.</summary>
    [Parameter]
    public SwitchParameter IncludeTables { get; set; }

    /// <summary>Render headers and footers.</summary>
    [Parameter]
    public SwitchParameter IncludeHeaderFooters { get; set; }

    /// <summary>Render document notes.</summary>
    [Parameter]
    public SwitchParameter IncludeNotes { get; set; }

    /// <summary>Maximum number of system font families to discover.</summary>
    [Parameter]
    [ValidateRange(1, int.MaxValue)]
    public int? MaximumSystemFontFamilies { get; set; }

    /// <summary>Allow embedding fonts discovered on the current system.</summary>
    [Parameter]
    public SwitchParameter AllowSystemFontEmbedding { get; set; }

    /// <summary>Allow embedding fonts referenced by the RTF document.</summary>
    [Parameter]
    public SwitchParameter AllowDocumentFontEmbedding { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new RtfToPdfOptions();
        if (PdfOptions != null) options.PdfOptions = PdfOptions;
        SetBoundSwitch(nameof(IncludeHiddenText), IncludeHiddenText, value => options.IncludeHiddenText = value);
        SetBoundSwitch(nameof(IncludeImages), IncludeImages, value => options.IncludeImages = value);
        if (DefaultImageWidth.HasValue) options.DefaultImageWidth = DefaultImageWidth.Value;
        if (DefaultImageHeight.HasValue) options.DefaultImageHeight = DefaultImageHeight.Value;
        SetBoundSwitch(nameof(IncludeMetadata), IncludeMetadata, value => options.IncludeMetadata = value);
        SetBoundSwitch(nameof(IncludeTables), IncludeTables, value => options.IncludeTables = value);
        SetBoundSwitch(nameof(IncludeHeaderFooters), IncludeHeaderFooters, value => options.IncludeHeaderFooters = value);
        SetBoundSwitch(nameof(IncludeNotes), IncludeNotes, value => options.IncludeNotes = value);
        if (MaximumSystemFontFamilies.HasValue) options.MaximumSystemFontFamilies = MaximumSystemFontFamilies.Value;
        SetBoundSwitch(nameof(AllowSystemFontEmbedding), AllowSystemFontEmbedding, value => options.ResourcePolicy.AllowSystemFontEmbedding = value);
        SetBoundSwitch(nameof(AllowDocumentFontEmbedding), AllowDocumentFontEmbedding, value => options.ResourcePolicy.AllowDocumentFontEmbedding = value);
        WriteObject(options);
    }

    private void SetBoundSwitch(string name, SwitchParameter value, System.Action<bool> setter) {
        if (MyInvocation.BoundParameters.ContainsKey(name)) setter(value.IsPresent);
    }
}
