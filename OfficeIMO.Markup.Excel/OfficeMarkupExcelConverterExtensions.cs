using OfficeIMO.Excel;
using OfficeIMO.Core.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Markup.Excel;

/// <summary>Converts typed OfficeIMO Markup into editable Excel documents.</summary>
public static class OfficeMarkupExcelConverterExtensions {
    /// <summary>Converts markup into a detached document owned by the caller.</summary>
    public static ExcelDocument ToExcel(this OfficeMarkupDocument document, MarkupExcelExportOptions? options = null,
        CancellationToken cancellationToken = default) => document.ToExcel(options, out _, cancellationToken);

    /// <summary>Converts markup and retains mapping and native validation evidence.</summary>
    public static ExcelDocument ToExcel(this OfficeMarkupDocument document, MarkupExcelExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        cancellationToken.ThrowIfCancellationRequested();
        MarkupExcelExportOptions resolved = options ?? new MarkupExcelExportOptions();
        IReadOnlyList<OfficeMarkupDiagnostic> diagnostics = CollectDiagnostics(document, resolved);
        ExcelDocument value = new OfficeMarkupExcelExporter().Convert(document, resolved);
        var result = new OfficeMarkupConversionResult<ExcelDocument>(value, diagnostics);
        report = result.Report;
        try {
            cancellationToken.ThrowIfCancellationRequested();
            if (!result.Succeeded) throw new OfficeConversionException("Markup conversion failed validation.", report);
            if (resolved.FailOnLoss) report.ThrowIfLoss();
            return result.Value;
        } catch { result.Value.Dispose(); throw; }
    }

    /// <summary>Converts markup asynchronously and returns the document with its completed report.</summary>
    public static Task<OfficeConversionResult<ExcelDocument>> ToExcelAsync(this OfficeMarkupDocument document,
        MarkupExcelExportOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeConversionOutput.RunAsync(() => (document.ToExcel(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);

    /// <summary>Writes converted Excel output and returns its report.</summary>
    public static OfficeConversionReport SaveAsExcel(
        this OfficeMarkupDocument document, string path, MarkupExcelExportOptions? options = null, ExcelSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        using ExcelDocument value = document.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionFailures.Run(report, () => value.Save(path, saveOptions ?? CreateDefaultSaveOptions()));
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Writes converted Excel output and returns its report.</summary>
    public static OfficeConversionReport SaveAsExcel(
        this OfficeMarkupDocument document, Stream stream, MarkupExcelExportOptions? options = null, ExcelSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        using ExcelDocument value = document.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionFailures.Run(report, () => value.Save(stream, saveOptions ?? CreateDefaultSaveOptions()));
        return report;
    }

    /// <summary>Asynchronously writes converted Excel output and returns its report.</summary>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(
        this OfficeMarkupDocument document, string path, MarkupExcelExportOptions? options = null, ExcelSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        using ExcelDocument value = document.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(path, saveOptions ?? CreateDefaultSaveOptions(), cancellationToken)).ConfigureAwait(false);
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Asynchronously writes converted Excel output and returns its report.</summary>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(
        this OfficeMarkupDocument document, Stream stream, MarkupExcelExportOptions? options = null, ExcelSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        using ExcelDocument value = document.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        await OfficeConversionFailures.RunAsync(report, () => value.SaveAsync(stream, saveOptions ?? CreateDefaultSaveOptions(), cancellationToken)).ConfigureAwait(false);
        return report;
    }

    private static ExcelSaveOptions CreateDefaultSaveOptions() => new ExcelSaveOptions {
        SafePreflight = true,
        ValidateOpenXml = true,
        SafeRepairDefinedNames = true
    };

    private static IReadOnlyList<OfficeMarkupDiagnostic> CollectDiagnostics(
        OfficeMarkupDocument document,
        MarkupExcelExportOptions options) {
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
        block,
        OfficeConversionLossKind.Omission);
}
