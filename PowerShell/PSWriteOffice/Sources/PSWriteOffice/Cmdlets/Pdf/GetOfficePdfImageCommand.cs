using System.IO;
using System.Management.Automation;
using OfficeIMO.Pdf;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Gets or extracts image resources from a PDF.</summary>
/// <example>
///   <summary>Extract images from selected pages.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$proof = @(
///     Get-OfficePdfImage -Path .\Examples\Documents\Report.pdf -PageRange '1-2'
///     Get-OfficePdfImage -Path .\Examples\Documents\Report.pdf -OutputDirectory .\Examples\Documents\PdfImages -BaseName 'report-image'
/// )
/// $proof</code>
///   <para>Returns image metadata or writes extracted images to disk.</para>
/// </example>
[Cmdlet(VerbsCommon.Get, "OfficePdfImage")]
[OutputType(typeof(PdfExtractedImage), typeof(FileInfo))]
public sealed class GetOfficePdfImageCommand : PSCmdlet
{
    /// <summary>PDF file path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Optional page ranges such as 1-3,5.</summary>
    [Parameter]
    public string? PageRange { get; set; }

    /// <summary>Optional directory where images should be written.</summary>
    [Parameter]
    public string? OutputDirectory { get; set; }

    /// <summary>Base file name used when extracting images to disk.</summary>
    [Parameter]
    public string BaseName { get; set; } = "image";

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful password authentication, explicitly ignore owner-imposed extraction restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        var inputPath = PdfCommandUtilities.ResolvePath(this, Path);
        var document = PdfDocument.Load(
            inputPath,
            PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent));
        var images = string.IsNullOrWhiteSpace(PageRange)
            ? document.Images.Extract()
            : document.Images.Extract(PageRange!);
        if (!string.IsNullOrWhiteSpace(OutputDirectory))
        {
            var outputDirectory = PdfCommandUtilities.ResolvePath(this, OutputDirectory!);
            PdfCommandUtilities.EnsureOutputDirectory(outputDirectory);
            int savedIndex = 0;
            foreach (var image in images)
            {
                if (!image.IsImageFile || string.IsNullOrWhiteSpace(image.FileExtension))
                {
                    continue;
                }

                savedIndex++;
                string extension = image.FileExtension!.TrimStart('.');
                string fileName = $"{BaseName}-{savedIndex:D4}.{extension}";
                string outputPath = AtomicFileWriter.WriteUnique(outputDirectory, fileName, image.Bytes);
                WriteObject(new FileInfo(outputPath));
            }

            return;
        }

        foreach (var image in images)
        {
            WriteObject(image);
        }
    }
}
