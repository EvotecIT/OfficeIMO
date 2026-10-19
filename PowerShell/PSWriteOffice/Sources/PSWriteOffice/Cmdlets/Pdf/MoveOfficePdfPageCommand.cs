using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Moves selected pages before another page and writes a new PDF.</summary>
/// <example>
///   <summary>Move appendix pages to the end.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Move-OfficePdfPage -Path .\Examples\Documents\Report.pdf -PageRange '2-3' -BeforePage 6 -OutputPath .\Examples\Documents\Report-Reordered.pdf
///     Get-OfficePdfInfo -Path .\Examples\Documents\Report-Reordered.pdf | Select-Object PageCount
/// )
/// $proof</code>
///   <para>Moves selected pages before a target one-based page number and writes a new PDF.</para>
/// </example>
[Cmdlet(VerbsCommon.Move, "OfficePdfPage", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class MoveOfficePdfPageCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Page ranges such as 1-3,5.</summary>
    [Parameter(Mandatory = true)]
    public string PageRange { get; set; } = string.Empty;

    /// <summary>One-based page number before which selected pages are inserted. Use page count + 1 to move to the end.</summary>
    [Parameter(Mandatory = true)]
    public int BeforePage { get; set; }

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
        if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Write reordered PDF pages")) {
            return;
        }

        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfDocument.Load(
                PdfCommandUtilities.ResolvePath(this, Path),
                PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent))
            .Pages.Move(BeforePage, PageRange).Save(outputPath).RequireSuccess();
        WritePassThru(new FileInfo(outputPath));
    }
}
