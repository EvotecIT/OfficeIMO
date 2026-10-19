using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates a typed image for a PDF table cell.</summary>
/// <example>
///   <summary>Add a linked logo to a typed PDF table cell.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$logo = New-OfficePdfTableCellImage -Path .\logo.png -Width 28 -Height 28 -LinkUri 'https://example.com'
/// $cell = New-OfficePdfTableCell -Text 'Portal' -Image $logo</code>
///   <para>The image remains a native PDF table-cell visual and may carry its own link.</para>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePdfTableCellImage")]
[Alias("PdfTableCellImage")]
[OutputType(typeof(PdfTableCellImage))]
public sealed class NewOfficePdfTableCellImageCommand : PSCmdlet
{
    /// <summary>Raster image path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("ImagePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Rendered width in PDF points.</summary>
    [Parameter(Mandatory = true)]
    [ValidateRange(0.1D, double.MaxValue)]
    public double Width { get; set; }

    /// <summary>Rendered height in PDF points.</summary>
    [Parameter(Mandatory = true)]
    [ValidateRange(0.1D, double.MaxValue)]
    public double Height { get; set; }

    /// <summary>Optional absolute or catalog-base-relative URI linked from the image.</summary>
    [Parameter]
    public string? LinkUri { get; set; }

    /// <summary>Accessible annotation text for the image link.</summary>
    [Parameter]
    public string? LinkContents { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var bytes = PdfImageDocumentSource.FromFile(PdfCommandUtilities.ResolvePath(this, Path)).GetBytes();
        WriteObject(new PdfTableCellImage(bytes, Width, Height, linkUri: LinkUri, linkContents: LinkContents));
    }
}
