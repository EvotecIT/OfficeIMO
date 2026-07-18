using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Fluent image-export builder for an Excel range.
    /// </summary>
    public sealed class ExcelRangeImageExportBuilder : OfficeImageExportBuilder<ExcelRangeImageExportBuilder, ExcelImageExportOptions> {
        internal ExcelRangeImageExportBuilder(ExcelRange range, ExcelImageExportOptions? options = null)
            : base(
                options?.Clone() ?? new ExcelImageExportOptions(),
                (format, effective, cancellationToken) => range.ExportImage(format, effective, cancellationToken)) {
        }

        /// <summary>Enables or disables worksheet gridline rendering.</summary>
        public ExcelRangeImageExportBuilder WithGridlines(bool show = true) {
            Options.ShowGridlines = show;
            return this;
        }

        /// <summary>Disables worksheet gridline rendering.</summary>
        public ExcelRangeImageExportBuilder WithoutGridlines() => WithGridlines(false);

        /// <summary>Includes or excludes hidden rows and columns.</summary>
        public ExcelRangeImageExportBuilder IncludeHidden(bool include = true) {
            Options.IncludeHidden = include;
            return this;
        }

        /// <summary>Includes or excludes worksheet images.</summary>
        public ExcelRangeImageExportBuilder IncludeImages(bool include = true) {
            Options.IncludeImages = include;
            return this;
        }

        /// <summary>Includes or excludes worksheet charts.</summary>
        public ExcelRangeImageExportBuilder IncludeCharts(bool include = true) {
            Options.IncludeCharts = include;
            return this;
        }

        /// <summary>Includes or excludes supported drawing objects.</summary>
        public ExcelRangeImageExportBuilder IncludeDrawingObjects(bool include = true) {
            Options.IncludeDrawingObjects = include;
            return this;
        }

        /// <summary>Includes or excludes supported conditional-formatting visuals.</summary>
        public ExcelRangeImageExportBuilder IncludeConditionalFormatting(bool include = true) {
            Options.IncludeConditionalFormatting = include;
            return this;
        }

        /// <summary>Sets the date used to evaluate relative conditional-formatting rules.</summary>
        public ExcelRangeImageExportBuilder WithConditionalFormattingDate(DateTime date) {
            Options.ConditionalFormattingDate = date;
            return this;
        }

        /// <summary>Enables or disables visible cell comment bodies.</summary>
        public ExcelRangeImageExportBuilder ShowComments(bool show = true) {
            Options.ShowCommentBodies = show;
            return this;
        }
    }

    /// <summary>
    /// Fluent image-export builder for an Excel worksheet.
    /// </summary>
    public sealed class ExcelWorksheetImageExportBuilder : OfficeImageExportBuilder<ExcelWorksheetImageExportBuilder, ExcelWorksheetImageExportOptions> {
        internal ExcelWorksheetImageExportBuilder(ExcelSheet sheet, ExcelWorksheetImageExportOptions? options = null)
            : base(
                options?.CloneWorksheet() ?? new ExcelWorksheetImageExportOptions(),
                (format, effective, cancellationToken) => sheet.ExportImage(format, effective, cancellationToken)) {
        }

        /// <summary>Exports an explicit A1 range instead of the worksheet used range.</summary>
        public ExcelWorksheetImageExportBuilder ForRange(string range) {
            if (string.IsNullOrWhiteSpace(range)) {
                throw new ArgumentException("Worksheet image export range cannot be null or whitespace.", nameof(range));
            }

            Options.Range = range;
            return this;
        }

        /// <summary>Uses the worksheet print area when one is configured.</summary>
        public ExcelWorksheetImageExportBuilder UsePrintArea(bool use = true) {
            Options.UsePrintArea = use;
            return this;
        }

        /// <summary>Enables or disables worksheet gridline rendering.</summary>
        public ExcelWorksheetImageExportBuilder WithGridlines(bool show = true) {
            Options.ShowGridlines = show;
            return this;
        }

        /// <summary>Disables worksheet gridline rendering.</summary>
        public ExcelWorksheetImageExportBuilder WithoutGridlines() => WithGridlines(false);
    }

    /// <summary>
    /// Fluent image-export builder for selected worksheets in an Excel workbook.
    /// </summary>
    public sealed class ExcelWorkbookImageExportBuilder : OfficeImageExportBatchBuilder<ExcelWorkbookImageExportBuilder, ExcelWorkbookImageExportOptions> {
        internal ExcelWorkbookImageExportBuilder(ExcelDocument workbook, ExcelWorkbookImageExportOptions? options = null)
            : base(
                options?.CloneWorkbook() ?? new ExcelWorkbookImageExportOptions(),
                workbook.ExportImages,
                (format, effective, consumer, cancellationToken) =>
                    workbook.ExportImages(format, consumer, effective, cancellationToken)) {
        }

        /// <summary>Exports only the named worksheets.</summary>
        public ExcelWorkbookImageExportBuilder ForSheets(params string[] sheetNames) => ForSheets((IEnumerable<string>)sheetNames);

        /// <summary>Exports only the named worksheets.</summary>
        public ExcelWorkbookImageExportBuilder ForSheets(IEnumerable<string> sheetNames) {
            if (sheetNames == null) {
                throw new ArgumentNullException(nameof(sheetNames));
            }

            var names = new List<string>();
            foreach (string sheetName in sheetNames) {
                if (string.IsNullOrWhiteSpace(sheetName)) {
                    throw new ArgumentException("Worksheet names cannot contain null or whitespace entries.", nameof(sheetNames));
                }

                names.Add(sheetName);
            }

            Options.SheetNames = names.AsReadOnly();
            return this;
        }

        /// <summary>Uses each worksheet's print area when one is configured.</summary>
        public ExcelWorkbookImageExportBuilder UsePrintAreas(bool use = true) {
            Options.UseWorksheetPrintAreas = use;
            return this;
        }

        /// <summary>Includes or excludes hidden and very hidden worksheets when exporting all workbook sheets.</summary>
        public ExcelWorkbookImageExportBuilder IncludeHiddenSheets(bool include = true) {
            Options.IncludeHiddenSheets = include;
            return this;
        }

        /// <summary>Splits worksheet exports by manual page breaks.</summary>
        public ExcelWorkbookImageExportBuilder SplitByManualPageBreaks(bool split = true) {
            Options.SplitWorksheetsByManualPageBreaks = split;
            return this;
        }

        /// <summary>Enables or disables worksheet gridline rendering.</summary>
        public ExcelWorkbookImageExportBuilder WithGridlines(bool show = true) {
            Options.ShowGridlines = show;
            return this;
        }

        /// <summary>Disables worksheet gridline rendering.</summary>
        public ExcelWorkbookImageExportBuilder WithoutGridlines() => WithGridlines(false);

        /// <summary>Includes or excludes hidden rows and columns.</summary>
        public ExcelWorkbookImageExportBuilder IncludeHidden(bool include = true) {
            Options.IncludeHidden = include;
            return this;
        }

        /// <summary>Includes or excludes worksheet images.</summary>
        public ExcelWorkbookImageExportBuilder IncludeImages(bool include = true) {
            Options.IncludeImages = include;
            return this;
        }

        /// <summary>Includes or excludes worksheet charts.</summary>
        public ExcelWorkbookImageExportBuilder IncludeCharts(bool include = true) {
            Options.IncludeCharts = include;
            return this;
        }

        /// <summary>Includes or excludes supported drawing objects.</summary>
        public ExcelWorkbookImageExportBuilder IncludeDrawingObjects(bool include = true) {
            Options.IncludeDrawingObjects = include;
            return this;
        }

        /// <summary>Includes or excludes supported conditional-formatting visuals.</summary>
        public ExcelWorkbookImageExportBuilder IncludeConditionalFormatting(bool include = true) {
            Options.IncludeConditionalFormatting = include;
            return this;
        }

        /// <summary>Sets the date used to evaluate relative conditional-formatting rules.</summary>
        public ExcelWorkbookImageExportBuilder WithConditionalFormattingDate(DateTime date) {
            Options.ConditionalFormattingDate = date;
            return this;
        }

        /// <summary>Enables or disables visible cell comment bodies.</summary>
        public ExcelWorkbookImageExportBuilder ShowComments(bool show = true) {
            Options.ShowCommentBodies = show;
            return this;
        }
    }

    public sealed partial class ExcelRange {
        /// <summary>
        /// Starts a fluent image export for this range.
        /// </summary>
        public ExcelRangeImageExportBuilder ToImage() => new ExcelRangeImageExportBuilder(this);

        /// <summary>Starts a fluent image export using a cloned options snapshot.</summary>
        public ExcelRangeImageExportBuilder ToImage(ExcelImageExportOptions options) =>
            new ExcelRangeImageExportBuilder(this, options ?? throw new ArgumentNullException(nameof(options)));
    }

    public partial class ExcelSheet {
        /// <summary>
        /// Starts a fluent image export for this worksheet.
        /// </summary>
        public ExcelWorksheetImageExportBuilder ToImage() => new ExcelWorksheetImageExportBuilder(this);

        /// <summary>Starts a fluent image export using a cloned options snapshot.</summary>
        public ExcelWorksheetImageExportBuilder ToImage(ExcelWorksheetImageExportOptions options) =>
            new ExcelWorksheetImageExportBuilder(this, options ?? throw new ArgumentNullException(nameof(options)));
    }

    public partial class ExcelDocument {
        /// <summary>
        /// Starts a fluent image export for selected worksheets in this workbook.
        /// </summary>
        public ExcelWorkbookImageExportBuilder ToImages() => new ExcelWorkbookImageExportBuilder(this);

        /// <summary>Starts a fluent batch export using a cloned options snapshot.</summary>
        public ExcelWorkbookImageExportBuilder ToImages(ExcelWorkbookImageExportOptions options) =>
            new ExcelWorkbookImageExportBuilder(this, options ?? throw new ArgumentNullException(nameof(options)));
    }
}
