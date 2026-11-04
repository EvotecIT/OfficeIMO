using System.Management.Automation;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Rtf;

/// <summary>Creates discoverable RTF-to-PDF conversion options for Export-OfficeDocumentPdf.</summary>
/// <example>
///   <summary>Include document structure and bound system-font discovery.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficeRtfPdfOptions -NoImages -NoTables -NoHeaderFooters -MaximumSystemFontFamilies 32
/// Export-OfficeDocumentPdf -InputPath .\Report.rtf -Path .\Report.pdf -RtfOptions $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficeRtfPdfOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(RtfPdfExportOptions))]
public sealed class NewOfficeRtfPdfOptionsCommand : PSCmdlet {
    /// <summary>Underlying low-level OfficeIMO PDF options.</summary>
    [Parameter(ParameterSetName = "Default")]
    public OfficeIMO.Pdf.PdfCreateOptions? PdfOptions { get; set; }

    /// <summary>Include text marked hidden.</summary>
    [Parameter]
    public SwitchParameter IncludeHiddenText { get; set; }

    /// <summary>Render images.</summary>
    [Parameter]
    public SwitchParameter NoImages { get; set; }

    /// <summary>Fallback image width in PDF points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? DefaultImageWidth { get; set; }

    /// <summary>Fallback image height in PDF points.</summary>
    [Parameter]
    [OfficeLengthTransformation]
    public OfficeLength? DefaultImageHeight { get; set; }

    /// <summary>Copy document metadata into the PDF.</summary>
    [Parameter]
    public SwitchParameter NoMetadata { get; set; }

    /// <summary>Render tables.</summary>
    [Parameter]
    public SwitchParameter NoTables { get; set; }

    /// <summary>Render headers and footers.</summary>
    [Parameter]
    public SwitchParameter NoHeaderFooters { get; set; }

    /// <summary>Render document notes.</summary>
    [Parameter]
    public SwitchParameter NoNotes { get; set; }

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
        if (DefaultImageWidth?.ToPoints() <= 0 || DefaultImageHeight?.ToPoints() <= 0)
            throw new PSArgumentOutOfRangeException(nameof(DefaultImageWidth), DefaultImageWidth, "Fallback image dimensions must be positive.");
        var options = new RtfPdfExportOptions();
        if (PdfOptions != null) options.PdfOptions = PdfOptions;
        SetBoundSwitch(nameof(IncludeHiddenText), IncludeHiddenText, value => options.IncludeHiddenText = value);
        SetBoundSwitch(nameof(NoImages), NoImages, value => options.IncludeImages = !value);
        if (DefaultImageWidth.HasValue) options.DefaultImageWidth = DefaultImageWidth.Value.ToPoints();
        if (DefaultImageHeight.HasValue) options.DefaultImageHeight = DefaultImageHeight.Value.ToPoints();
        SetBoundSwitch(nameof(NoMetadata), NoMetadata, value => options.IncludeMetadata = !value);
        SetBoundSwitch(nameof(NoTables), NoTables, value => options.IncludeTables = !value);
        SetBoundSwitch(nameof(NoHeaderFooters), NoHeaderFooters, value => options.IncludeHeaderFooters = !value);
        SetBoundSwitch(nameof(NoNotes), NoNotes, value => options.IncludeNotes = !value);
        if (MaximumSystemFontFamilies.HasValue) options.MaximumSystemFontFamilies = MaximumSystemFontFamilies.Value;
        SetBoundSwitch(nameof(AllowSystemFontEmbedding), AllowSystemFontEmbedding, value => options.ResourcePolicy.AllowSystemFontEmbedding = value);
        SetBoundSwitch(nameof(AllowDocumentFontEmbedding), AllowDocumentFontEmbedding, value => options.ResourcePolicy.AllowDocumentFontEmbedding = value);
        WriteObject(options);
    }

    private void SetBoundSwitch(string name, SwitchParameter value, System.Action<bool> setter) {
        if (MyInvocation.BoundParameters.ContainsKey(name)) setter(value.IsPresent);
    }
}
