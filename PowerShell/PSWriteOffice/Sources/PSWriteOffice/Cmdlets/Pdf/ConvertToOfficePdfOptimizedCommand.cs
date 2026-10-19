using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Applies lossless PDF optimization actions and writes a new PDF.</summary>
/// <example>
///   <summary>Optimize a PDF with lossless stream compression.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficePdfOptimized -Path .\Report.pdf -OutputPath .\Report-Optimized.pdf</code>
///   <para>Writes a smaller PDF when safe lossless optimization actions can reduce the file size.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfOptimized", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
[OutputType(typeof(PdfOptimizationActionResult))]
public sealed class ConvertToOfficePdfOptimizedCommand : PSCmdlet
{
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

    /// <summary>After successful password authentication, explicitly ignore owner-imposed modification restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Skip Flate compression of unfiltered streams.</summary>
    [Parameter]
    public SwitchParameter NoCompressStreams { get; set; }

    /// <summary>Keep orphaned indirect PDF objects instead of pruning objects unreachable from the catalog.</summary>
    [Parameter]
    public SwitchParameter KeepUnreferencedObjects { get; set; }

    /// <summary>Keep byte-identical stream objects instead of rewriting duplicate references to one object.</summary>
    [Parameter]
    public SwitchParameter KeepDuplicateStreams { get; set; }

    /// <summary>Write the optimized candidate even when it is not smaller than the source PDF.</summary>
    [Parameter]
    public SwitchParameter AllowLarger { get; set; }

    /// <summary>Minimum unfiltered stream size considered for compression.</summary>
    [Parameter]
    public int MinimumStreamCompressionBytes { get; set; } = 128;

    /// <summary>Return the optimization action report instead of the output file.</summary>
    [Parameter]
    public SwitchParameter PassThruReport { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        string inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        string outputPath = PdfCommandUtilities.ResolvePath(this, OutputPath);
        if (!ShouldProcess(outputPath, "Write optimized PDF"))
        {
            return;
        }

        var options = new PdfOptimizationOptions
        {
            CompressUnfilteredStreams = !NoCompressStreams.IsPresent,
            RemoveUnreferencedObjects = !KeepUnreferencedObjects.IsPresent,
            DeduplicateIdenticalStreams = !KeepDuplicateStreams.IsPresent,
            KeepOriginalWhenNotSmaller = !AllowLarger.IsPresent,
            MinimumStreamCompressionBytes = MinimumStreamCompressionBytes
        };

        PdfCommandUtilities.EnsureDirectory(outputPath);
        PdfOptimizationActionResult result = PdfDocument.Load(
            inputPath,
            PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent)).Optimization.Apply(options);
        result.ToDocument().Save(outputPath).RequireSuccess();
        WriteObject(PassThruReport.IsPresent ? result : new FileInfo(outputPath));
    }
}
