using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets a simple or fully composed running PDF header.</summary>
/// <example>
///   <summary>Add a running report header.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfHeader.pdf {
///     Set-OfficePdfHeader -Text 'Service Review' -Align Right -FontSize 9
///     Add-OfficePdfHeading -Text 'Service Review'
///     Add-OfficePdfParagraph -Text 'The header repeats on generated pages.'
/// }</code>
///   <para>Sets header text for the generated PDF.</para>
/// </example>
/// <example>
///   <summary>Compose styled default, first-page, and even-page headers.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\RichHeader.pdf {
///     Set-OfficePdfHeader -Compose {
///         param($header)
///         $label = New-OfficeTextRun -Text 'Service report ' -Bold | ConvertTo-OfficePdfTextRun
///         $pageStyle = New-OfficeTextRun -Italic | ConvertTo-OfficePdfTextRun
///         $null = $header.Text({
///             param($text)
///             $null = $text.Run($label).CurrentPage($pageStyle)
///         })
///         $null = $header.FirstPageText('Service report cover')
///         $null = $header.EvenPagesZones('Service report', $null, 'Page {page}/{pages}')
///     }
///     Add-OfficePdfParagraph -Text 'Generated report body.'
/// }</code>
///   <para>The native composer owns rich runs, page tokens, zones, images, shapes, and page variants.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfHeader", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfHeader")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfHeaderCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Compatibility parameter. Page composition is supported only inside New-OfficePdf.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Header text. Supports {page} and {pages}.</summary>
    [Parameter(Position = 0)]
    public string? Text { get; set; }

    /// <summary>
    /// Advanced header composer. The script receives a <see cref="PdfHeaderBuilder"/> and can configure
    /// default, first-page, and even-page text, zones, images, shapes, rich text, and page tokens.
    /// </summary>
    [Parameter]
    public ScriptBlock? Compose { get; set; }

    /// <summary>Header alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Center;

    /// <summary>Header font size in PDF points.</summary>
    [Parameter]
    public double? FontSize { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var hasText = MyInvocation.BoundParameters.ContainsKey(nameof(Text));
        if (Compose != null && hasText)
        {
            throw new PSArgumentException("Use either -Text or -Compose, not both.");
        }

        if (Compose == null && (!hasText || Text == null))
        {
            throw new PSArgumentException("Provide -Text or -Compose.");
        }

        var compose = Compose?.GetNewClosure();
        var document = PdfCommandUtilities.ComposePage(this, Document, ParameterSetName, ParameterSetDocument, page => page.Header(header =>
        {
            ApplyAlignment(header);
            if (FontSize.HasValue)
            {
                header.FontSize(FontSize.Value);
            }
            if (compose != null)
            {
                compose.Invoke(header);
            }
            else
            {
                header.Text(Text!);
            }
        }));

        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }

    private void ApplyAlignment(PdfHeaderBuilder header)
    {
        switch (Align)
        {
            case PdfAlign.Right:
                header.AlignRight();
                break;
            case PdfAlign.Left:
                header.AlignLeft();
                break;
            default:
                header.AlignCenter();
                break;
        }
    }
}
