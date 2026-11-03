using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

public static partial class ExcelIWorkConverter {
    /// <summary>Opens and converts a Numbers file into the normal editable Excel model.</summary>
    public static ExcelDocument ConvertNumbersToExcel(string path,
        IWorkReadOptions? readOptions = null, IWorkConversionOptions? conversionOptions = null) =>
        ConvertNumbersToExcelResult(path, readOptions, conversionOptions).RequireCompleteEditableReconstruction();

    /// <summary>Opens and converts a Numbers stream into the normal editable Excel model.</summary>
    public static ExcelDocument ConvertNumbersToExcel(Stream stream,
        IWorkReadOptions? readOptions = null, IWorkConversionOptions? conversionOptions = null) =>
        ConvertNumbersToExcelResult(stream, readOptions, conversionOptions).RequireCompleteEditableReconstruction();

    /// <summary>Opens and converts a Numbers file, returning the workbook with source evidence.</summary>
    public static NumbersToExcelResult ConvertNumbersToExcelResult(string path,
        IWorkReadOptions? readOptions = null, IWorkConversionOptions? conversionOptions = null) {
        ArgumentNullException.ThrowIfNull(path);
        return ProjectNumbers(
            IWorkSourceDocument.Open(path, IWorkDocumentKind.Numbers, readOptions), conversionOptions);
    }

    /// <summary>Opens and converts a Numbers stream, returning the workbook with source evidence.</summary>
    public static NumbersToExcelResult ConvertNumbersToExcelResult(Stream stream,
        IWorkReadOptions? readOptions = null, IWorkConversionOptions? conversionOptions = null) {
        ArgumentNullException.ThrowIfNull(stream);
        return ProjectNumbers(
            IWorkSourceDocument.Open(stream, IWorkDocumentKind.Numbers, readOptions), conversionOptions);
    }

    /// <summary>Converts an opened Numbers source into the normal editable Excel model.</summary>
    public static ExcelDocument ToExcelDocument(this IWorkSourceDocument source,
        IWorkConversionOptions? options = null) =>
        ToExcelDocumentResult(source, options).RequireCompleteEditableReconstruction();

    /// <summary>Converts an opened Numbers source and returns the workbook with its projection and loss report.</summary>
    public static NumbersToExcelResult ToExcelDocumentResult(this IWorkSourceDocument source,
        IWorkConversionOptions? options = null) {
        ArgumentNullException.ThrowIfNull(source);
        return ProjectNumbers(source, options);
    }
}
