using System.Data;
using OfficeIMO.Excel;
using PSWriteOffice.Services;
using PSWriteOffice.Services.Excel;

namespace PSWriteOffice.Cmdlets.Excel;

public sealed partial class ExportOfficeExcelCommand
{
    private bool TryWriteDataReaderPackageDirectly(
        IDataReader reader,
        string resolvedPath,
        bool preserveWorkbook,
        ExcelTableStyle style,
        ExcelColumnFormatPlan? columnFormatPlan)
    {
        if (!CanWriteDataReaderPackageDirectly(preserveWorkbook, columnFormatPlan))
        {
            return false;
        }

        var options = new ExcelWriterOptions
        {
            SheetName = WorksheetName,
            IncludeHeaders = !NoHeader.IsPresent,
            CreateTable = !NoTable.IsPresent,
            TableName = string.IsNullOrWhiteSpace(TableName) ? "ReaderData" : TableName,
            TableStyle = style,
            IncludeAutoFilter = !NoAutoFilter.IsPresent,
            IncludeCellReferences = false,
            UseSharedStrings = false,
            DateSystem = DateSystem ?? ExcelDateSystem.NineteenHundred
        };
        var result = ExcelDocumentService.WriteDataReaderPackage(
            resolvedPath,
            new PowerShellNormalizingDataReader(reader, CreateNormalizerOptions()),
            options,
            overwrite: !NoClobber.IsPresent);

        if (!string.IsNullOrWhiteSpace(result.Range))
        {
            WriteVerbose($"Exported data reader to {result.SheetName}!{result.Range} through the streaming package writer.");
        }

        if (Open.IsPresent)
        {
            FileOpenService.Open(resolvedPath);
        }

        WritePassThru(resolvedPath);
        return true;
    }

    private bool CanWriteDataReaderPackageDirectly(
        bool preserveWorkbook,
        ExcelColumnFormatPlan? columnFormatPlan)
    {
        return SaveOptions == null && Theme == null && !preserveWorkbook &&
            !AppendsToNamedTable &&
            StartRow == 1 &&
            StartColumn == 1 &&
            (!NoHeader.IsPresent || NoTable.IsPresent) &&
            string.IsNullOrWhiteSpace(Title) &&
            ExcludeProperty is not { Length: > 0 } &&
            columnFormatPlan == null &&
            !AutoFit.IsPresent &&
            !AutoFitFormattedColumn.IsPresent &&
            !BoldTopRow.IsPresent &&
            !FreezeTopRow.IsPresent &&
            !FreezeFirstColumn.IsPresent &&
            !ShowFirstColumn.IsPresent &&
            !ShowLastColumn.IsPresent &&
            !NoRowStripes.IsPresent &&
            !ShowColumnStripes.IsPresent &&
            !HasWorkbookProperties();
    }

    private bool HasWorkbookProperties() => DocumentProperty is { Count: > 0 };
}
