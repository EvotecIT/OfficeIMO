using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel;

/// <summary>
/// Fluent batch image export for worksheet print areas and manual page-break segments.
/// </summary>
public sealed class ExcelWorksheetImageBatchExportBuilder
    : OfficeImageExportBatchBuilder<ExcelWorksheetImageBatchExportBuilder, ExcelWorksheetImageExportOptions> {
    internal ExcelWorksheetImageBatchExportBuilder(
        ExcelSheet sheet,
        ExcelWorksheetImageExportOptions? options = null)
        : base(
            options?.CloneWorksheet() ?? new ExcelWorksheetImageExportOptions(),
            sheet.ExportImages,
            (format, effective, consumer, cancellationToken) =>
                sheet.ExportImages(format, consumer, effective, cancellationToken)) {
    }

    /// <summary>Exports an explicit A1 range instead of the worksheet used range.</summary>
    public ExcelWorksheetImageBatchExportBuilder ForRange(string range) {
        if (string.IsNullOrWhiteSpace(range)) {
            throw new ArgumentException("Worksheet image export range cannot be null or whitespace.", nameof(range));
        }
        Options.Range = range;
        return this;
    }

    /// <summary>Uses every configured worksheet print-area segment.</summary>
    public ExcelWorksheetImageBatchExportBuilder UsePrintArea(bool use = true) {
        Options.UsePrintArea = use;
        return this;
    }

    /// <summary>Splits the selected range at manual row and column page breaks.</summary>
    public ExcelWorksheetImageBatchExportBuilder SplitByManualPageBreaks(bool split = true) {
        Options.SplitByManualPageBreaks = split;
        return this;
    }

    /// <summary>Enables or disables worksheet gridline rendering.</summary>
    public ExcelWorksheetImageBatchExportBuilder WithGridlines(bool show = true) {
        Options.ShowGridlines = show;
        return this;
    }

    /// <summary>Includes or excludes hidden rows and columns.</summary>
    public ExcelWorksheetImageBatchExportBuilder IncludeHidden(bool include = true) {
        Options.IncludeHidden = include;
        return this;
    }

    /// <summary>Includes or excludes worksheet images.</summary>
    public ExcelWorksheetImageBatchExportBuilder IncludeImages(bool include = true) {
        Options.IncludeImages = include;
        return this;
    }

    /// <summary>Includes or excludes worksheet charts.</summary>
    public ExcelWorksheetImageBatchExportBuilder IncludeCharts(bool include = true) {
        Options.IncludeCharts = include;
        return this;
    }

    /// <summary>Includes or excludes supported drawing objects.</summary>
    public ExcelWorksheetImageBatchExportBuilder IncludeDrawingObjects(bool include = true) {
        Options.IncludeDrawingObjects = include;
        return this;
    }

    /// <summary>Includes or excludes supported conditional-formatting visuals.</summary>
    public ExcelWorksheetImageBatchExportBuilder IncludeConditionalFormatting(bool include = true) {
        Options.IncludeConditionalFormatting = include;
        return this;
    }

    /// <summary>Enables or disables visible cell comment bodies.</summary>
    public ExcelWorksheetImageBatchExportBuilder ShowComments(bool show = true) {
        Options.ShowCommentBodies = show;
        return this;
    }

    /// <summary>Enables or disables automatic hyperlink styling hints.</summary>
    public ExcelWorksheetImageBatchExportBuilder ShowHyperlinkHints(bool show = true) {
        Options.ShowHyperlinkHints = show;
        return this;
    }
}

public partial class ExcelSheet {
    /// <summary>
    /// Starts a fluent batch export for print-area segments or manual page-break regions.
    /// </summary>
    public ExcelWorksheetImageBatchExportBuilder ToImages() => new ExcelWorksheetImageBatchExportBuilder(this);

    /// <summary>Starts a fluent worksheet batch export using a cloned options snapshot.</summary>
    public ExcelWorksheetImageBatchExportBuilder ToImages(ExcelWorksheetImageExportOptions options) =>
        new ExcelWorksheetImageBatchExportBuilder(
            this,
            options ?? throw new ArgumentNullException(nameof(options)));
}
