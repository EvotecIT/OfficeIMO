using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Removes selected pages from a PDF and writes a new PDF.</summary>
/// <example>
///   <summary>Remove draft pages from a PDF.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Remove-OfficePdfPage -Path .\Examples\Documents\Report.pdf -PageRange '4-5' -OutputPath .\Examples\Documents\Report-Clean.pdf
///     Get-OfficePdfPreflight -Path .\Examples\Documents\Report-Clean.pdf
/// )
/// $proof</code>
///   <para>Deletes selected pages and preflights the rewritten PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Remove, "OfficePdfPage", SupportsShouldProcess = true, ConfirmImpact = ConfirmImpact.Medium)]
[OutputType(typeof(FileInfo))]
public sealed class RemoveOfficePdfPageCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
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
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write PDF without selected pages")) {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfDocument.Load(
                PdfCommandUtilities.ResolvePath(this, Path),
                PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent))
            .Pages.Delete(PageRange).Save(outputPath).RequireSuccess();
        WritePassThru(new FileInfo(outputPath));
    }
}
