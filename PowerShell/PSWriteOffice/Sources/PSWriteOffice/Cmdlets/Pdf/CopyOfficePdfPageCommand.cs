using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Copies selected PDF pages into a new PDF.</summary>
/// <example>
///   <summary>Extract a page range into a new PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Copy-OfficePdfPage -Path .\Examples\Documents\Report.pdf -PageRange '1-2,5' -OutputPath .\Examples\Documents\ExecutivePages.pdf
///     Get-OfficePdfInfo -Path .\Examples\Documents\ExecutivePages.pdf | Select-Object PageCount
/// )
/// $proof</code>
///   <para>Copies selected pages and inspects the resulting PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Copy, "OfficePdfPage", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class CopyOfficePdfPageCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Page ranges such as 1-3,5.</summary>
    [Parameter(Mandatory = true)]
    public string PageRange { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed assembly restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        var outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write copied PDF pages")) {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfDocument.Load(
                PdfCommandUtilities.ResolvePath(this, Path),
                PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent))
            .Pages.Extract(PageRange).Save(outputPath).RequireSuccess();
        WritePassThru(new FileInfo(outputPath));
    }
}
