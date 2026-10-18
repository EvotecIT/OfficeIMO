using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.PowerPoint.Pdf;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Reconstructs a PowerPoint presentation from a PDF.</summary>
/// <para>Defaults to editable content and also supports visual-page, editable-table, and hybrid reconstruction through OfficeIMO options.</para>
/// <example>
///   <summary>Convert a PDF to an editable PowerPoint presentation.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficePdfPowerPoint -Path .\Briefing.pdf -OutputPath .\Briefing.pptx</code>
///   <para>Writes a PPTX deck using OfficeIMO's richest safe editable projection.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfPowerPoint", SupportsShouldProcess = true)]
[Alias("ConvertTo-PdfPowerPoint")]
[OutputType(typeof(FileInfo))]
[OutputType(typeof(PdfPowerPointConversionReport))]
public sealed class ConvertToOfficePdfPowerPointCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output PPTX path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("OutPath")]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful authentication, explicitly ignore owner-imposed extraction restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Advanced OfficeIMO PDF-to-PowerPoint reconstruction options.</summary>
    [Parameter]
    public PdfToPowerPointOptions? Options { get; set; }

    /// <summary>Overwrite an existing output file.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Open the converted presentation after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Return the detailed conversion report instead of file information.</summary>
    [Parameter]
    public SwitchParameter PassThruReport { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        string? outputPath = null;
        var outputOperation = false;
        try
        {
            var inputPath = PdfCommandUtilities.ResolveExistingFilePath(this, Path);
            outputOperation = true;
            outputPath = PdfCommandUtilities.ResolveOutputFilePath(this, OutputPath, ".pptx", Force.IsPresent);
            if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Convert PDF to PowerPoint presentation"))
            {
                return;
            }

            outputOperation = false;
            var document = PdfCommandUtilities.LoadDocument(
                inputPath,
                PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent));
            outputOperation = true;
            PdfCommandUtilities.EnsureDirectory(outputPath);
            var outputResult = document.SaveAsPowerPoint(outputPath, Options).RequireSuccess();
            var report = outputResult.Report
                ?? throw new InvalidOperationException("PDF-to-PowerPoint conversion completed without a conversion report.");
            if (Open.IsPresent)
            {
                FileOpenService.Open(outputPath);
            }

            WriteObject(PassThruReport.IsPresent ? report : new FileInfo(outputPath));
        }
        catch (Exception exception)
        {
            WriteError(PdfCommandUtilities.CreateConversionErrorRecord(
                exception,
                "ConvertToOfficePdfPowerPointFailed",
                outputOperation ? outputPath ?? OutputPath : Path,
                outputOperation));
        }
    }
}
