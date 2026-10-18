using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Word.Pdf;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Reconstructs editable Word content from a PDF.</summary>
/// <para>Uses OfficeIMO's semantic PDF reader to recover supported headings, paragraphs, lists, tables, links, images, metadata, and page breaks.</para>
/// <example>
///   <summary>Convert a PDF to an editable Word document.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficePdfWord -Path .\Report.pdf -OutputPath .\Report.docx</code>
///   <para>Writes a DOCX document and returns its file information.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfWord", SupportsShouldProcess = true)]
[Alias("ConvertTo-PdfWord")]
[OutputType(typeof(FileInfo))]
[OutputType(typeof(PdfWordConversionReport))]
public sealed class ConvertToOfficePdfWordCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output DOCX path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("OutPath")]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful authentication, explicitly ignore owner-imposed extraction restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Advanced OfficeIMO PDF-to-Word reconstruction options.</summary>
    [Parameter]
    public PdfToWordOptions? Options { get; set; }

    /// <summary>Overwrite an existing output file.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Open the converted document after saving.</summary>
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
            outputPath = PdfCommandUtilities.ResolveOutputFilePath(this, OutputPath, ".docx", Force.IsPresent);
            if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Convert PDF to editable Word document"))
            {
                return;
            }

            outputOperation = false;
            var document = PdfCommandUtilities.LoadDocument(
                inputPath,
                PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent));
            outputOperation = true;
            PdfCommandUtilities.EnsureDirectory(outputPath);
            var outputResult = document.SaveAsWord(outputPath, Options).RequireSuccess();
            var report = outputResult.Report
                ?? throw new InvalidOperationException("PDF-to-Word conversion completed without a conversion report.");
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
                "ConvertToOfficePdfWordFailed",
                outputOperation ? outputPath ?? OutputPath : Path,
                outputOperation));
        }
    }
}
