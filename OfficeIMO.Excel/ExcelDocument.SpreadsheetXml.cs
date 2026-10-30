using OfficeIMO.Excel.SpreadsheetXml;
using OfficeIMO.Core.Internal;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private long _sourceMaxCharactersInPart;
    /// <summary>
    /// Imports a bounded XML Spreadsheet 2003 file into the normal editable workbook model.
    /// Unsupported features are reported, never evaluated. The caller owns the returned workbook.
    /// </summary>
    public static OfficeConversionResult<ExcelDocument, ExcelSpreadsheetXmlReport> ImportSpreadsheetXml(
        string filePath, ExcelSpreadsheetXmlOptions? options = null) {
        if (filePath == null) throw new ArgumentNullException(nameof(filePath));
        using var input = File.OpenRead(filePath);
        return ImportSpreadsheetXml(input, options);
    }

    /// <summary>Imports XML Spreadsheet bytes with operation-level fidelity diagnostics.</summary>
    public static OfficeConversionResult<ExcelDocument, ExcelSpreadsheetXmlReport> ImportSpreadsheetXml(
        byte[] input, ExcelSpreadsheetXmlOptions? options = null) {
        if (input == null) throw new ArgumentNullException(nameof(input));
        using var stream = new MemoryStream(input, writable: false);
        return ImportSpreadsheetXml(stream, options);
    }

    /// <summary>
    /// Imports the remaining bytes of a readable stream. The stream remains open and its position advances.
    /// Use RequireNoLoss on the result when omissions must reject the workflow; dispose its Value even on rejection.
    /// </summary>
    public static OfficeConversionResult<ExcelDocument, ExcelSpreadsheetXmlReport> ImportSpreadsheetXml(
        Stream input, ExcelSpreadsheetXmlOptions? options = null) =>
        SpreadsheetXmlReader.Read(input, (options ?? new ExcelSpreadsheetXmlOptions()).Snapshot());

    /// <summary>
    /// Exports the current workbook to XML Spreadsheet 2003 data interchange bytes and reports unsupported content.
    /// Does not change the associated file path. XML is a separate explicit route, not an alias for XLSX Save.
    /// </summary>
    public OfficeConversionResult<byte[], ExcelSpreadsheetXmlReport> ExportSpreadsheetXml(
        ExcelSpreadsheetXmlOptions? options = null) {
        var resolved = (options ?? new ExcelSpreadsheetXmlOptions()).Snapshot();
        resolved.CancellationToken.ThrowIfCancellationRequested();
        EnsureLegacyXlsSaveDoesNotDropImportedContent(options: null);
        bool readOnly = _spreadSheetDocument.FileOpenAccess == FileAccess.Read;
        byte[] bytes = readOnly
            ? Locking.ExecuteRead(EnsureLock(), () => CaptureSpreadsheetXmlSnapshot(resolved, readOnly: true))
            : Locking.ExecuteWrite(EnsureLock(), () => CaptureSpreadsheetXmlSnapshot(resolved, readOnly: false));
        return SpreadsheetXmlWriter.Write(bytes, resolved, _sourceMaxCharactersInPart);
    }

    private byte[] CaptureSpreadsheetXmlSnapshot(ExcelSpreadsheetXmlOptions options, bool readOnly) {
        // Flush deferred edits through the normal owner, without running save-time calculation callbacks
        // or clearing caches. Clone the original package so omitted parts can still be diagnosed.
        if (!readOnly) PrepareWorkbookForSave(options: null, applyCalculationPolicy: false);
        var carriageReturns = CollectLoadedCarriageReturnParts(WorkbookPartRoot);
        using var snapshot = new OfficeBoundedMemoryStream(options.MaxInputBytes);
        using (_spreadSheetDocument.Clone(snapshot, false)) { }
        if (carriageReturns.Count == 0) return snapshot.ToArray();
        using var rewritten = new OfficeBoundedMemoryStream(options.MaxInputBytes);
        RewritePackageWithPreservedCarriageReturns(snapshot, rewritten, carriageReturns);
        return rewritten.ToArray();
    }
}
