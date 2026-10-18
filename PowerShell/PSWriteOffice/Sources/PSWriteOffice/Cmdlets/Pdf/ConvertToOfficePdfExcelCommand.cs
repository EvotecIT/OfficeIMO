using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel.Pdf;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Pdf;

/// <summary>Extracts detected PDF tables into an editable Excel workbook.</summary>
/// <para>Uses OfficeIMO's logical table reconstruction, including typed numeric, Boolean, date, and percentage columns plus multi-page table continuation handling.</para>
/// <example>
///   <summary>Convert detected PDF tables to Excel.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficePdfExcel -Path .\Statement.pdf -OutputPath .\Statement.xlsx</code>
///   <para>Writes an XLSX workbook with one table per detected logical table.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficePdfExcel", SupportsShouldProcess = true)]
[Alias("ConvertTo-PdfExcel")]
[OutputType(typeof(FileInfo))]
[OutputType(typeof(PdfExcelTableImportReport))]
public sealed class ConvertToOfficePdfExcelCommand : PSCmdlet
{
    /// <summary>Input PDF path.</summary>
    [Parameter(Mandatory = true, Position = 0, ValueFromPipeline = true)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Output XLSX path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("OutPath")]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Password used to authenticate an encrypted PDF.</summary>
    [Parameter]
    public string? Password { get; set; }

    /// <summary>After successful authentication, explicitly ignore owner-imposed extraction restrictions.</summary>
    [Parameter]
    public SwitchParameter IgnorePermissionRestrictions { get; set; }

    /// <summary>Advanced OfficeIMO PDF-table-to-Excel options.</summary>
    [Parameter]
    public PdfTablesToExcelOptions? Options { get; set; }

    /// <summary>Overwrite an existing output file.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Open the converted workbook after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Return the detailed table import report instead of file information.</summary>
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
            outputPath = PdfCommandUtilities.ResolveOutputFilePath(this, OutputPath, ".xlsx", Force.IsPresent);
            if (!PdfCommandUtilities.ShouldWrite(this, outputPath, "Convert PDF tables to editable Excel workbook"))
            {
                return;
            }

            outputOperation = false;
            var document = PdfCommandUtilities.LoadDocument(
                inputPath,
                PdfCommandUtilities.CreateReadOptions(Password, IgnorePermissionRestrictions.IsPresent));
            outputOperation = true;
            PdfCommandUtilities.EnsureDirectory(outputPath);
            var outputResult = document.SaveTablesAsExcel(outputPath, Options).RequireSuccess();
            var report = outputResult.Report
                ?? throw new InvalidOperationException("PDF-to-Excel conversion completed without a conversion report.");
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
                "ConvertToOfficePdfExcelFailed",
                outputOperation ? outputPath ?? OutputPath : Path,
                outputOperation));
        }
    }
}
