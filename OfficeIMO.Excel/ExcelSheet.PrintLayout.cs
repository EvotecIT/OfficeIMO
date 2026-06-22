namespace OfficeIMO.Excel {
    /// <summary>
    /// Common print layout workflows for worksheets.
    /// </summary>
    public enum ExcelPrintLayoutPreset {
        /// <summary>Portrait worksheet defaults with normal margins.</summary>
        Worksheet,
        /// <summary>Landscape report layout, one page wide, repeated first header row, narrow margins.</summary>
        Report,
        /// <summary>Landscape dashboard layout, fit to one page, narrow margins.</summary>
        Dashboard,
        /// <summary>Landscape data-table layout, one page wide, repeated first header row.</summary>
        DataTable
    }

    /// <summary>
    /// Options for applying a worksheet print layout preset.
    /// Nullable values override the selected preset only when provided.
    /// </summary>
    public sealed class ExcelPrintLayoutOptions {
        /// <summary>Preset to apply.</summary>
        public ExcelPrintLayoutPreset Preset { get; set; } = ExcelPrintLayoutPreset.Report;

        /// <summary>Optional print area in A1 notation.</summary>
        public string? PrintArea { get; set; }

        /// <summary>Optional orientation override.</summary>
        public ExcelPageOrientation? Orientation { get; set; }

        /// <summary>Optional margin preset override.</summary>
        public ExcelMarginPreset? Margins { get; set; }

        /// <summary>Optional pages-wide fit override.</summary>
        public uint? FitToWidth { get; set; }

        /// <summary>Optional pages-tall fit override. Use 0 for unlimited height.</summary>
        public uint? FitToHeight { get; set; }

        /// <summary>Optional manual scale percentage override.</summary>
        public uint? Scale { get; set; }

        /// <summary>Optional page order override.</summary>
        public ExcelPageOrder? PageOrder { get; set; }

        /// <summary>Optional first repeated print-title row.</summary>
        public int? RepeatFirstRow { get; set; }

        /// <summary>Optional last repeated print-title row.</summary>
        public int? RepeatLastRow { get; set; }

        /// <summary>Optional first repeated print-title column.</summary>
        public int? RepeatFirstColumn { get; set; }

        /// <summary>Optional last repeated print-title column.</summary>
        public int? RepeatLastColumn { get; set; }

        /// <summary>When true, preset default print-title rows are not applied.</summary>
        public bool SuppressPresetPrintTitles { get; set; }
    }

    public partial class ExcelSheet {
        /// <summary>
        /// Applies a reusable worksheet print layout preset with optional overrides.
        /// </summary>
        /// <param name="preset">Preset to apply.</param>
        /// <param name="printArea">Optional print area in A1 notation.</param>
        /// <returns>The worksheet for fluent chaining.</returns>
        public ExcelSheet ApplyPrintLayoutPreset(ExcelPrintLayoutPreset preset, string? printArea = null) {
            return ApplyPrintLayout(new ExcelPrintLayoutOptions {
                Preset = preset,
                PrintArea = printArea,
            });
        }

        /// <summary>
        /// Applies reusable worksheet print layout settings.
        /// </summary>
        /// <param name="options">Print layout options.</param>
        /// <returns>The worksheet for fluent chaining.</returns>
        public ExcelSheet ApplyPrintLayout(ExcelPrintLayoutOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));

            var preset = ResolvePrintLayoutPreset(options.Preset);
            var orientation = options.Orientation ?? preset.Orientation;
            var margins = options.Margins ?? preset.Margins;
            var fitToWidth = options.FitToWidth ?? preset.FitToWidth;
            var fitToHeight = options.FitToHeight ?? preset.FitToHeight;
            var scale = options.Scale ?? preset.Scale;
            var pageOrder = options.PageOrder ?? preset.PageOrder;
            var repeatFirstRow = options.RepeatFirstRow ?? (options.SuppressPresetPrintTitles ? null : preset.RepeatFirstRow);
            var repeatLastRow = options.RepeatLastRow ?? (options.SuppressPresetPrintTitles ? null : preset.RepeatLastRow);
            var repeatFirstColumn = options.RepeatFirstColumn ?? preset.RepeatFirstColumn;
            var repeatLastColumn = options.RepeatLastColumn ?? preset.RepeatLastColumn;

            SetOrientation(orientation);
            SetMarginsPreset(margins);
            SetPageSetup(fitToWidth, fitToHeight, scale, pageOrder);

            if (!string.IsNullOrWhiteSpace(options.PrintArea)) {
                _excelDocument.SetPrintArea(this, options.PrintArea!, save: false);
            }

            if (repeatFirstRow.HasValue || repeatLastRow.HasValue || repeatFirstColumn.HasValue || repeatLastColumn.HasValue) {
                _excelDocument.SetPrintTitles(
                    this,
                    repeatFirstRow,
                    repeatLastRow,
                    repeatFirstColumn,
                    repeatLastColumn,
                    save: false);
            }

            WorksheetRoot.Save();
            return this;
        }

        private static ResolvedPrintLayoutPreset ResolvePrintLayoutPreset(ExcelPrintLayoutPreset preset) {
            switch (preset) {
                case ExcelPrintLayoutPreset.Worksheet:
                    return new ResolvedPrintLayoutPreset(
                        ExcelPageOrientation.Portrait,
                        ExcelMarginPreset.Normal,
                        fitToWidth: null,
                        fitToHeight: null,
                        scale: 100,
                        ExcelPageOrder.DownThenOver,
                        repeatFirstRow: null,
                        repeatLastRow: null,
                        repeatFirstColumn: null,
                        repeatLastColumn: null);
                case ExcelPrintLayoutPreset.Dashboard:
                    return new ResolvedPrintLayoutPreset(
                        ExcelPageOrientation.Landscape,
                        ExcelMarginPreset.Narrow,
                        fitToWidth: 1,
                        fitToHeight: 1,
                        scale: null,
                        ExcelPageOrder.OverThenDown,
                        repeatFirstRow: null,
                        repeatLastRow: null,
                        repeatFirstColumn: null,
                        repeatLastColumn: null);
                case ExcelPrintLayoutPreset.DataTable:
                    return new ResolvedPrintLayoutPreset(
                        ExcelPageOrientation.Landscape,
                        ExcelMarginPreset.Normal,
                        fitToWidth: 1,
                        fitToHeight: 0,
                        scale: null,
                        ExcelPageOrder.DownThenOver,
                        repeatFirstRow: 1,
                        repeatLastRow: 1,
                        repeatFirstColumn: null,
                        repeatLastColumn: null);
                default:
                    return new ResolvedPrintLayoutPreset(
                        ExcelPageOrientation.Landscape,
                        ExcelMarginPreset.Narrow,
                        fitToWidth: 1,
                        fitToHeight: 0,
                        scale: null,
                        ExcelPageOrder.DownThenOver,
                        repeatFirstRow: 1,
                        repeatLastRow: 1,
                        repeatFirstColumn: null,
                        repeatLastColumn: null);
            }
        }

        private sealed class ResolvedPrintLayoutPreset {
            internal ResolvedPrintLayoutPreset(
                ExcelPageOrientation orientation,
                ExcelMarginPreset margins,
                uint? fitToWidth,
                uint? fitToHeight,
                uint? scale,
                ExcelPageOrder pageOrder,
                int? repeatFirstRow,
                int? repeatLastRow,
                int? repeatFirstColumn,
                int? repeatLastColumn) {
                Orientation = orientation;
                Margins = margins;
                FitToWidth = fitToWidth;
                FitToHeight = fitToHeight;
                Scale = scale;
                PageOrder = pageOrder;
                RepeatFirstRow = repeatFirstRow;
                RepeatLastRow = repeatLastRow;
                RepeatFirstColumn = repeatFirstColumn;
                RepeatLastColumn = repeatLastColumn;
            }

            internal ExcelPageOrientation Orientation { get; }
            internal ExcelMarginPreset Margins { get; }
            internal uint? FitToWidth { get; }
            internal uint? FitToHeight { get; }
            internal uint? Scale { get; }
            internal ExcelPageOrder PageOrder { get; }
            internal int? RepeatFirstRow { get; }
            internal int? RepeatLastRow { get; }
            internal int? RepeatFirstColumn { get; }
            internal int? RepeatLastColumn { get; }
        }
    }
}
