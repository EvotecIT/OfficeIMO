using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Removes PDF annotations matching friendly filters.</summary>
/// <example>
///   <summary>Remove text annotations from the first page.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>Remove-OfficePdfAnnotation -Path .\Reviewed.pdf -OutputPath .\Clean.pdf -PageNumber 1 -Subtype Text -Confirm:$false</code>
/// </example>
[Cmdlet(VerbsCommon.Remove, "OfficePdfAnnotation", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
[OutputType(typeof(PdfAnnotationEditResult))]
public sealed class RemoveOfficePdfAnnotationCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output PDF path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed annotation-modification restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Specific annotation object number to remove.</summary>
    [Parameter]
    public int? ObjectNumber { get; set; }

    /// <summary>Optional one-based page number filter.</summary>
    [Parameter]
    public int? PageNumber { get; set; }

    /// <summary>Optional annotation subtype filter such as Text, Link, Widget, or FreeText.</summary>
    [Parameter]
    public string? Subtype { get; set; }

    /// <summary>Keep popup annotations linked from matching annotations through /Popup.</summary>
    [Parameter]
    public SwitchParameter KeepPopups { get; set; }

    /// <summary>Return the annotation edit result instead of the output file.</summary>
    [Parameter]
    public SwitchParameter PassThruReport { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        string inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        string outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!ShouldProcess(outputPath, "Remove PDF annotations")) {
            return;
        }

        var options = new PdfAnnotationRemovalOptions {
            ObjectNumber = ObjectNumber,
            PageNumber = PageNumber,
            Subtype = Subtype,
            RemoveMatchingPopups = !KeepPopups.IsPresent
        };

        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfAnnotationEditResult result = PdfDocument
            .Load(inputPath, PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent))
            .Annotations.Remove(options);
        result.ToDocument().Save(outputPath).RequireSuccess();
        if (PassThruReport.IsPresent) {
            WriteObject(result);
            return;
        }

        WritePassThru(new FileInfo(outputPath));
    }
}
