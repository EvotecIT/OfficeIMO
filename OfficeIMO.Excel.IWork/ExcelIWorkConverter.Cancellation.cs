using System.Threading;
using OfficeIMO.IWork;

namespace OfficeIMO.Excel.IWork;

public static partial class ExcelIWorkConverter {
    /// <summary>Converts a Numbers file or directory bundle with cancellation and returns the destination document.</summary>
    public static ExcelDocument ConvertNumbersToExcel(string path,
        IWorkReadOptions? readOptions, IWorkConversionOptions? conversionOptions,
        CancellationToken cancellationToken) =>
        ConvertNumbersToExcelResult(path, readOptions, conversionOptions, cancellationToken).RequireCompleteEditableReconstruction();

    /// <summary>Converts a Numbers file or directory bundle with cancellation and retains source evidence.</summary>
    /// <remarks>Cancellation governs loading, semantic projection, and destination construction. Saving is a separate owner operation.</remarks>
    public static NumbersToExcelResult ConvertNumbersToExcelResult(string path,
        IWorkReadOptions? readOptions, IWorkConversionOptions? conversionOptions,
        CancellationToken cancellationToken) {
        if (path == null) throw new ArgumentNullException(nameof(path));
        cancellationToken.ThrowIfCancellationRequested();
        return ProjectNumbers(IWorkSourceDocument.Open(path, IWorkDocumentKind.Numbers,
            readOptions, cancellationToken), conversionOptions);
    }

    /// <summary>Converts a Numbers caller-owned ZIP stream with cancellation and returns the destination document.</summary>
    public static ExcelDocument ConvertNumbersToExcel(Stream stream,
        IWorkReadOptions? readOptions, IWorkConversionOptions? conversionOptions,
        CancellationToken cancellationToken) =>
        ConvertNumbersToExcelResult(stream, readOptions, conversionOptions, cancellationToken).RequireCompleteEditableReconstruction();

    /// <summary>Converts a Numbers caller-owned ZIP stream with cancellation and retains source evidence.</summary>
    /// <remarks>Cancellation governs loading, semantic projection, and destination construction. Saving is a separate owner operation.</remarks>
    public static NumbersToExcelResult ConvertNumbersToExcelResult(Stream stream,
        IWorkReadOptions? readOptions, IWorkConversionOptions? conversionOptions,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        cancellationToken.ThrowIfCancellationRequested();
        return ProjectNumbers(IWorkSourceDocument.Open(stream, IWorkDocumentKind.Numbers,
            readOptions, cancellationToken), conversionOptions);
    }
}
