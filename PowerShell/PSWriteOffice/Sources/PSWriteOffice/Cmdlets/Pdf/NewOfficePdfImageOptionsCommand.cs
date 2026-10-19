using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Cmdlets.Imaging;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Creates discoverable thumbnail and rendering settings for Export-OfficePdfImage.</summary>
/// <example>
///   <summary>Create compact PDF thumbnails with bounded output dimensions.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$options = New-OfficePdfImageOptions -ThumbnailMaxDimension 320 -MaximumOutputWidth 640
/// Export-OfficePdfImage -Path .\Report.pdf -OutputPath .\Thumbnails -Options $options</code>
/// </example>
[Cmdlet(VerbsCommon.New, "OfficePdfImageOptions")]
[OutputType(typeof(PdfImageExportOptions))]
public sealed class NewOfficePdfImageOptionsCommand : OfficeImageOptionsCommandBase<PdfImageExportOptions> {
    /// <summary>Maximum thumbnail width or height.</summary>
    [Parameter] [ValidateRange(1, int.MaxValue)] public int? ThumbnailMaxDimension { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var options = new PdfImageExportOptions();
        ApplyCommon(options);
        if (ThumbnailMaxDimension.HasValue) options.ThumbnailMaxDimension = ThumbnailMaxDimension.Value;
        WriteObject(options);
    }
}
