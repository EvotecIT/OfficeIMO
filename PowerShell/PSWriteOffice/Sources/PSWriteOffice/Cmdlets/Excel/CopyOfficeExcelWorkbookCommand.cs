using System.IO;
using System.Management.Automation;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

/// <summary>Copies a workbook package while preserving package parts.</summary>
/// <example>
///   <summary>Copy a workbook package and return the copied file.</summary>
///   <prefix>PS&gt; </prefix>
///   <code>$copy = Copy-OfficeExcelWorkbook -Path .\Template.xlsx -DestinationPath .\Report.xlsx -Force -PassThru
/// Test-OfficeExcelWorkbook -Path $copy.FullName -SkipOpenXmlValidation |
///     Select-Object Passed, WorksheetCount</code>
///   <para>Copies the workbook package and normalizes the workbook content type for the destination extension.</para>
/// </example>
[Cmdlet(VerbsCommon.Copy, "OfficeExcelWorkbook", SupportsShouldProcess = true)]
[Alias("ExcelWorkbookCopy", "ExcelPackageCopy")]
[OutputType(typeof(FileInfo))]
public sealed class CopyOfficeExcelWorkbookCommand : PSCmdlet {
    /// <summary>Source workbook or template package path.</summary>
    [Parameter(Mandatory = true, Position = 0)]
    [Alias("FilePath", "InputPath", "SourcePath")]
    public string Path { get; set; } = string.Empty;

    /// <summary>Destination workbook path.</summary>
    [Parameter(Mandatory = true, Position = 1)]
    [Alias("Destination", "OutputPath", "TargetPath")]
    public string DestinationPath { get; set; } = string.Empty;

    /// <summary>Replace an existing destination workbook.</summary>
    [Parameter]
    public SwitchParameter Force { get; set; }

    /// <summary>Emit a <see cref="FileInfo"/> for the copied workbook.</summary>
    [Parameter]
    public SwitchParameter PassThru { get; set; }

    /// <inheritdoc />
    protected override void ProcessRecord() {
        string sourcePath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(Path);
        string destinationPath = SessionState.Path.GetUnresolvedProviderPathFromPSPath(DestinationPath);

        if (!ShouldProcess(destinationPath, $"Copy workbook package from '{sourcePath}'")) {
            return;
        }

        ExcelDocumentService.CopyWorkbookPackage(sourcePath, destinationPath, Force.IsPresent);

        if (PassThru.IsPresent) {
            WriteObject(new FileInfo(destinationPath));
        }
    }
}