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

        var options = new ExcelTabularWriteOptions
        {
            SheetName = WorksheetName,
            IncludeHeaders = !NoHeader.IsPresent,
            CreateTable = !NoTable.IsPresent,
            TableName = TableName,
            TableStyle = style,
            IncludeAutoFilter = !NoAutoFilter.IsPresent,
            IncludeCellReferences = false,
            UseSharedStrings = false,
            DateSystem = string.IsNullOrWhiteSpace(DateSystem)
                ? ExcelDateSystem.NineteenHundred
                : ExcelDateSystemService.Resolve(DateSystem!, nameof(DateSystem))
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
        return !preserveWorkbook &&
            !AppendToTable.IsPresent &&
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
            !SafePreflight.IsPresent &&
            !SafeRepairDefinedNames.IsPresent &&
            !ValidateOpenXml.IsPresent &&
            !DisableFastPackageWriter.IsPresent &&
            !EvaluateFormulas.IsPresent &&
            !ClearCachedFormulaResults.IsPresent &&
            !MarkFormulasDirty.IsPresent &&
            !ForceFullCalculationOnOpen.IsPresent &&
            !HasWorkbookProperties();
    }

    private bool HasWorkbookProperties()
    {
        return !string.IsNullOrWhiteSpace(DocumentTitle) ||
            !string.IsNullOrWhiteSpace(Author) ||
            !string.IsNullOrWhiteSpace(Subject) ||
            !string.IsNullOrWhiteSpace(Keywords) ||
            !string.IsNullOrWhiteSpace(Description) ||
            !string.IsNullOrWhiteSpace(Category) ||
            !string.IsNullOrWhiteSpace(Company) ||
            !string.IsNullOrWhiteSpace(Manager) ||
            !string.IsNullOrWhiteSpace(ApplicationName) ||
            !string.IsNullOrWhiteSpace(LastModifiedBy);
    }
}
