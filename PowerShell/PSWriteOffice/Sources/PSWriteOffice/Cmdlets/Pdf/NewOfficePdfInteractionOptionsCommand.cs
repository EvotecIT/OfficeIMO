using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates native region limits and invisible-text policy for Get-OfficePdfInteractionMap.</summary>
/// <example><summary>Include an existing OCR text layer in the interaction map.</summary>
/// <code>$options = New-OfficePdfInteractionOptions -IncludeInvisibleText -MaxTextRegions 50000
/// Get-OfficePdfInteractionMap -Path ./Scanned.pdf -Options $options</code></example>
[Cmdlet(VerbsCommon.New, "OfficePdfInteractionOptions", DefaultParameterSetName = "Default")]
[OutputType(typeof(PdfPageInteractionOptions))]
public sealed class NewOfficePdfInteractionOptionsCommand : OfficeCmdlet {
    /// <summary>Maximum text-element regions emitted for a page.</summary>
    [Parameter(ParameterSetName = "Default"), ValidateRange(1, int.MaxValue)]
    public int? MaxTextRegions { get; set; }
    /// <summary>Maximum image-placement regions emitted for a page.</summary>
    [Parameter, ValidateRange(1, int.MaxValue)]
    public int? MaxImageRegions { get; set; }
    /// <summary>Include invisible extracted text such as an OCR layer.</summary>
    [Parameter]
    public SwitchParameter IncludeInvisibleText { get; set; }
    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PdfPageInteractionOptions { IncludeInvisibleText = IncludeInvisibleText.IsPresent };
        if (MaxTextRegions.HasValue) options.MaxTextRegions = MaxTextRegions.Value;
        if (MaxImageRegions.HasValue) options.MaxImageRegions = MaxImageRegions.Value;
        WriteObject(options);
    }
}
