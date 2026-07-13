using OfficeIMO.Excel;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Markup.Excel;

/// <summary>Typed OfficeIMO Markup to Excel conversion helpers.</summary>
public static class OfficeMarkupExcelConverterExtensions {
    /// <summary>Converts workbook-profile markup to an editable detached Excel document.</summary>
    public static ExcelDocument ToExcelDocument(this OfficeMarkupDocument document, MarkupToExcelOptions? options = null) {
        OfficeMarkupConversionResult<ExcelDocument> result = document.ToExcelDocumentResult(options);
        if (result.Succeeded) return result.Value;
        result.Value.Dispose();
        return result.RequireValue();
    }

    /// <summary>Converts workbook-profile markup and returns the Excel document with structured diagnostics.</summary>
    public static OfficeMarkupConversionResult<ExcelDocument> ToExcelDocumentResult(
        this OfficeMarkupDocument document,
        MarkupToExcelOptions? options = null) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        MarkupToExcelOptions resolved = options ?? new MarkupToExcelOptions();
        IReadOnlyList<OfficeMarkupDiagnostic> diagnostics = CollectDiagnostics(document, resolved);
        ExcelDocument value = new OfficeMarkupExcelExporter().Convert(document, resolved);
        return new OfficeMarkupConversionResult<ExcelDocument>(value, diagnostics);
    }

    /// <summary>Converts workbook-profile markup and saves it as an Excel workbook.</summary>
    public static OfficeMarkupConversionReport SaveAsExcel(
        this OfficeMarkupDocument document,
        string path,
        MarkupToExcelOptions? options = null,
        ExcelSaveOptions? saveOptions = null) {
        OfficeMarkupConversionResult<ExcelDocument> result = document.ToExcelDocumentResult(options);
        using (result.Value) result.RequireValue().Save(path, saveOptions ?? CreateDefaultSaveOptions());
        return result.Report;
    }

    /// <summary>Converts workbook-profile markup and writes it to a caller-owned stream.</summary>
    public static OfficeMarkupConversionReport SaveAsExcel(
        this OfficeMarkupDocument document,
        Stream stream,
        MarkupToExcelOptions? options = null,
        ExcelSaveOptions? saveOptions = null) {
        OfficeMarkupConversionResult<ExcelDocument> result = document.ToExcelDocumentResult(options);
        using (result.Value) result.RequireValue().Save(stream, saveOptions ?? CreateDefaultSaveOptions());
        return result.Report;
    }

    /// <summary>Converts workbook-profile markup and asynchronously saves it as an Excel workbook.</summary>
    public static async Task<OfficeMarkupConversionReport> SaveAsExcelAsync(
        this OfficeMarkupDocument document,
        string path,
        MarkupToExcelOptions? options = null,
        ExcelSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeMarkupConversionResult<ExcelDocument> result = document.ToExcelDocumentResult(options);
        using (result.Value) await result.RequireValue().SaveAsync(path, saveOptions ?? CreateDefaultSaveOptions(), cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    /// <summary>Converts workbook-profile markup and asynchronously writes it to a caller-owned stream.</summary>
    public static async Task<OfficeMarkupConversionReport> SaveAsExcelAsync(
        this OfficeMarkupDocument document,
        Stream stream,
        MarkupToExcelOptions? options = null,
        ExcelSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeMarkupConversionResult<ExcelDocument> result = document.ToExcelDocumentResult(options);
        using (result.Value) await result.RequireValue().SaveAsync(stream, saveOptions ?? CreateDefaultSaveOptions(), cancellationToken).ConfigureAwait(false);
        return result.Report;
    }

    private static ExcelSaveOptions CreateDefaultSaveOptions() => new ExcelSaveOptions {
        SafePreflight = true,
        ValidateOpenXml = true,
        SafeRepairDefinedNames = true
    };

    private static IReadOnlyList<OfficeMarkupDiagnostic> CollectDiagnostics(
        OfficeMarkupDocument document,
        MarkupToExcelOptions options) {
        var diagnostics = new List<OfficeMarkupDiagnostic>(OfficeMarkupValidator.Validate(document));
        foreach (OfficeMarkupBlock block in document.DescendantsAndSelf()) {
            bool commonText = block is OfficeMarkupHeadingBlock
                || block is OfficeMarkupParagraphBlock
                || block is OfficeMarkupListBlock
                || block is OfficeMarkupTableBlock;
            if (commonText && !options.IncludeMarkdownAsWorksheetText) {
                diagnostics.Add(Omitted(block));
            } else if (block is OfficeMarkupCodeBlock
                || block is OfficeMarkupImageBlock
                || block is OfficeMarkupDiagramBlock
                || block is OfficeMarkupExtensionBlock
                || block is OfficeMarkupRawMarkdownBlock) {
                diagnostics.Add(Omitted(block));
            }
        }

        return diagnostics.AsReadOnly();
    }

    private static OfficeMarkupDiagnostic Omitted(OfficeMarkupBlock block) => new OfficeMarkupDiagnostic(
        OfficeMarkupDiagnosticSeverity.Warning,
        $"{block.Kind} markup is omitted from Excel output.",
        block);
}
