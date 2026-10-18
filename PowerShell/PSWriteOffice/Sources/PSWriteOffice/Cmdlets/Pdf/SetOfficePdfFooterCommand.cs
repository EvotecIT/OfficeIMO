using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Sets a simple or fully composed running PDF footer.</summary>
/// <example>
///   <summary>Add page numbers to a generated PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\PdfFooter.pdf {
///     Set-OfficePdfFooter -Text 'Page {page} of {pages}' -Align Center -FontSize 8
///     Add-OfficePdfHeading -Text 'Report with footer'
///     Add-OfficePdfPageBreak
///     Add-OfficePdfParagraph -Text 'The footer includes generated page numbers.'
/// }</code>
///   <para>Uses page placeholders in a running footer.</para>
/// </example>
/// <example>
///   <summary>Compose a styled footer with page tokens.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>New-OfficePdf -Path .\Examples\Documents\RichFooter.pdf {
///     Set-OfficePdfFooter -Compose {
///         param($footer)
///         $label = New-OfficeTextRun -Text 'Confidential - page ' -Bold | ConvertTo-OfficePdfTextRun
///         $pageStyle = New-OfficeTextRun -Italic | ConvertTo-OfficePdfTextRun
///         $null = $footer.AlignRight().Text({
///             param($text)
///             $null = $text.Run($label).CurrentPage($pageStyle).Text(' of ').TotalPages($pageStyle)
///         })
///     }
///     Add-OfficePdfParagraph -Text 'Generated report body.'
/// }</code>
///   <para>Styled page tokens remain live until the document is rendered.</para>
/// </example>
[Cmdlet(VerbsCommon.Set, "OfficePdfFooter", DefaultParameterSetName = ParameterSetContext)]
[Alias("PdfFooter")]
[OutputType(typeof(PdfDocument))]
public sealed class SetOfficePdfFooterCommand : PSCmdlet
{
    private const string ParameterSetContext = "Context";
    private const string ParameterSetDocument = "Document";

    /// <summary>Compatibility parameter. Page composition is supported only inside New-OfficePdf.</summary>
    [Parameter(Mandatory = true, ValueFromPipeline = true, ParameterSetName = ParameterSetDocument)]
    public PdfDocument Document { get; set; } = null!;

    /// <summary>Footer text. Supports {page} and {pages}.</summary>
    [Parameter(Position = 0)]
    public string Text { get; set; } = "Page {page}/{pages}";

    /// <summary>
    /// Advanced footer composer. The script receives a <see cref="PdfFooterBuilder"/> and can configure
    /// default, first-page, and even-page text, zones, images, shapes, rich text, and page tokens.
    /// </summary>
    [Parameter]
    public ScriptBlock? Compose { get; set; }

    /// <summary>Footer alignment.</summary>
    [Parameter]
    public PdfAlign Align { get; set; } = PdfAlign.Center;

    /// <summary>Footer font size in PDF points.</summary>
    [Parameter]
    public double? FontSize { get; set; }

    /// <summary>Emit the updated document.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        if (Compose != null && MyInvocation.BoundParameters.ContainsKey(nameof(Text)))
        {
            throw new PSArgumentException("Use either -Text or -Compose, not both.");
        }

        var compose = Compose?.GetNewClosure();
        var document = PdfCommandUtilities.ComposePage(this, Document, ParameterSetName, ParameterSetDocument, page => page.Footer(footer =>
        {
            ApplyAlignment(footer);
            if (FontSize.HasValue)
            {
                footer.FontSize(FontSize.Value);
            }
            if (compose != null)
            {
                compose.Invoke(footer);
            }
            else
            {
                footer.Text(Text);
            }
        }));

        if (PassThru.IsPresent && document != null)
        {
            WriteObject(document);
        }
    }

    private void ApplyAlignment(PdfFooterBuilder footer)
    {
        switch (Align)
        {
            case PdfAlign.Right:
                footer.AlignRight();
                break;
            case PdfAlign.Left:
                footer.AlignLeft();
                break;
            default:
                footer.AlignCenter();
                break;
        }
    }
}
