using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Converts reusable Office text run specifications to native PDF text runs.</summary>
/// <remarks>
/// Use this adapter when an OfficeIMO PDF callback requires a native <see cref="PdfTextRun"/>,
/// such as a rich generated header or footer. Styling remains owned by <c>New-OfficeTextRun</c>.
/// </remarks>
/// <example>
///   <summary>Create native styled runs for a generated PDF header.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$label = New-OfficeTextRun -Text 'Service report ' -Bold -Color '#B42318' | ConvertTo-OfficePdfTextRun
/// $pageStyle = New-OfficeTextRun -Italic | ConvertTo-OfficePdfTextRun
/// Set-OfficePdfHeader -Compose {
///     param($header)
///     $null = $header.Text({
///         param($text)
///         $null = $text.Run($label).CurrentPage($pageStyle)
///     })
/// }</code>
///   <para>The cross-format run specification stays PowerShell-friendly while the callback receives the native PDF run it requires.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfTextRun")]
[Alias("PdfNativeTextRun")]
[OutputType(typeof(PdfTextRun))]
public sealed class ConvertToOfficePdfTextRunCommand : PSCmdlet
{
    /// <summary>One or more values accepted by New-OfficeTextRun, including run specifications and hashtables.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    public object Run { get; set; } = null!;

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        foreach (var run in PdfRichTextRunBuilder.ToTextRuns(Run))
        {
            WriteObject(run);
        }
    }
}
