using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Updates a single indirect PDF annotation.</summary>
[Cmdlet(VerbsCommon.Set, "OfficePdfAnnotation", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
[OutputType(typeof(PdfAnnotationEditResult))]
public sealed class SetOfficePdfAnnotationCommand : PSWriteOffice.Cmdlets.OfficeMutationCmdlet {
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

    /// <summary>Indirect annotation object number to update.</summary>
    [Parameter(Mandatory = true)]
    public int ObjectNumber { get; set; }

    /// <summary>Replacement annotation contents text.</summary>
    [Parameter]
    public string? Contents { get; set; }

    /// <summary>Replacement annotation title text.</summary>
    [Parameter]
    public string? Title { get; set; }

    /// <summary>Replacement annotation name.</summary>
    [Parameter]
    public string? Name { get; set; }

    /// <summary>Replacement annotation flags.</summary>
    [Parameter]
    public int? Flags { get; set; }

    /// <summary>Replacement annotation color. Named colors and hexadecimal values are accepted.</summary>
    [Parameter]
    [OfficeColorArgumentTransformation]
    [ArgumentCompleter(typeof(OfficeColorArgumentCompleter))]
    public string? Color { get; set; }

    /// <summary>Remove /A and /AA action dictionaries from the annotation.</summary>
    [Parameter]
    public SwitchParameter RemoveAction { get; set; }

    /// <summary>Return the annotation edit result instead of the output file.</summary>
    [Parameter]
    public SwitchParameter PassThruReport { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        string inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        string outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!ShouldProcess(outputPath, "Update PDF annotation")) {
            return;
        }

        PdfColor? color = PdfCommandUtilities.ParseColor(Color);
        var options = new PdfAnnotationUpdateOptions {
            Contents = Contents,
            Title = Title,
            Name = Name,
            Flags = Flags,
            Color = color.HasValue ? new[] { color.Value.R, color.Value.G, color.Value.B } : null,
            RemoveActions = RemoveAction.IsPresent
        };

        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfAnnotationEditResult result = PdfDocument
            .Load(inputPath, PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent))
            .Annotations.Update(ObjectNumber, options);
        result.ToDocument().Save(outputPath).RequireSuccess();
        if (PassThruReport.IsPresent) {
            WriteObject(result);
            return;
        }

        WritePassThru(new FileInfo(outputPath));
    }
}
