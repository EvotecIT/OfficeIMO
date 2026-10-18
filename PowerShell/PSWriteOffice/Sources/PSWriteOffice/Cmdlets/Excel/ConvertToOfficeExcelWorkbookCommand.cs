using System;
using System.IO;
using System.Management.Automation;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Pdf;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Converts Excel workbooks between supported .xls and .xlsx formats.</summary>
/// <para>Uses the OfficeIMO Excel normal load/save conversion path, including legacy XLS diagnostics and save preflight.</para>
/// <example>
///   <summary>Convert a legacy XLS file to XLSX.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficeExcelWorkbook -Path .\legacy.xls -OutputPath .\converted.xlsx -PassThru</code>
///   <para>Reads the .xls file and writes a .xlsx workbook.</para>
/// </example>
/// <example>
///   <summary>Convert an XLSX workbook to native XLS.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>ConvertTo-OfficeExcelWorkbook -Path .\report.xlsx -OutputPath .\report.xls -Force</code>
///   <para>Writes a supported native BIFF8 .xls workbook.</para>
/// </example>
[Cmdlet(VerbsData.ConvertTo, "OfficeExcelWorkbook", SupportsShouldProcess = true)]
[OutputType(typeof(FileInfo))]
public sealed class ConvertToOfficeExcelWorkbookCommand : PSCmdlet
{
    /// <summary>Source .xls or .xlsx file path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination .xls or .xlsx file path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("OutPath")]
    public string OutputPath { get; set; } = string.Empty;

    /// <summary>Overwrite an existing destination file.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Allow conversion when a legacy XLS source contains unsupported or preserve-only content.</summary>
    [Parameter]
    public SwitchParameter AllowLossyLegacyConversion { get; set; }

    /// <summary>Open the converted workbook after saving.</summary>
    [Parameter]
    public SwitchParameter Open { get; set; }

    /// <summary>Emit the saved file information.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord()
    {
        try
        {
            var sourcePath = ResolvePath(Path);
            var outputPath = ResolvePath(OutputPath);

            if (!File.Exists(sourcePath))
            {
                throw new FileNotFoundException($"File '{sourcePath}' was not found.", sourcePath);
            }

            if (File.Exists(outputPath) && !Force.IsPresent)
            {
                throw new IOException($"File '{outputPath}' already exists. Use -Force to overwrite it.");
            }

            var action = $"Convert Excel workbook to {System.IO.Path.GetExtension(outputPath)}";
            if (!ShouldProcess(outputPath, action))
            {
                return;
            }

            PdfCommandUtilities.EnsureDirectory(outputPath);
            ExcelDocument.Convert(sourcePath, outputPath, new ExcelDocumentConversionOptions
            {
                FileConflictPolicy = Force.IsPresent
                    ? OfficeConversionFileConflictPolicy.Replace
                    : OfficeConversionFileConflictPolicy.FailIfExists,
                LossPolicy = AllowLossyLegacyConversion.IsPresent
                    ? OfficeConversionLossPolicy.Allow
                    : OfficeConversionLossPolicy.Block
            });

            if (Open.IsPresent)
            {
                FileOpenService.Open(outputPath);
            }

            if (PassThru.IsPresent)
            {
                WriteObject(new FileInfo(outputPath));
            }
        }
        catch (Exception ex)
        {
            WriteError(new ErrorRecord(ex, "ConvertToOfficeExcelWorkbookFailed", ErrorCategory.InvalidOperation, Path));
        }
    }

    private string ResolvePath(string path)
    {
        return PdfCommandUtilities.ResolvePath(this, path);
    }
}
