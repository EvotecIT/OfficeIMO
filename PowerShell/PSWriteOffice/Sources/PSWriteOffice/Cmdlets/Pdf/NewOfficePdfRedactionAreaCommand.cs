using System.Management.Automation;
using OfficeIMO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using PSWriteOffice.Foundation;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates a redaction area for Remove-OfficePdfContent and Get-OfficePdfRedactionPlan -Area.</summary>
/// <example><summary>Use the native OfficePdfRedactionArea command.</summary><code>$area = New-OfficePdfRedactionArea -Page 1 -X 2cm -Y 3cm -Width 8cm -Height 2cm -Label 'Signature'
/// Remove-OfficePdfContent -Path .\Contracts\Signed.pdf -OutputPath .\Contracts\Signed-Redacted.pdf -Area $area</code></example>
[Cmdlet("New", "OfficePdfRedactionArea", DefaultParameterSetName = "Default")]
[OutputType(typeof(PdfRedactionArea))]
public sealed class NewOfficePdfRedactionAreaCommand : OfficeCmdlet {
    /// <summary>Page for this operation.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true, Position = 0)]
    public int Page { get; set; }
    /// <summary>X for this operation.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength X { get; set; }
    /// <summary>Y for this operation.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength Y { get; set; }
    /// <summary>Width for this operation.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength Width { get; set; }
    /// <summary>Height for this operation.</summary>
    [Parameter(ParameterSetName = "Default", Mandatory = true)]
    [OfficeLengthTransformation]
    public OfficeLength Height { get; set; }
    /// <summary>Label for this operation.</summary>
    [Parameter(ParameterSetName = "Default")]
    public string? Label { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() => WriteObject(new PdfRedactionArea(Page, X.ToPoints(), Y.ToPoints(), Width.ToPoints(), Height.ToPoints(), Label));
}
