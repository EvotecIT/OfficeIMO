namespace OfficeIMO.Excel {
    // Print and conditional formatting conveniences for ExcelReportBuilder.
    public sealed partial class ExcelReportBuilder {
        /// <summary>
        /// Applies sensible print defaults: gridlines off, fit to width, and optional print area.
        /// </summary>
        public ExcelReportBuilder PrintDefaults(bool showGridlines = false, uint fitToWidth = 1, uint fitToHeight = 0, string? printAreaA1 = null) {
            Sheet.SetGridlinesVisibleCore(showGridlines);
            Sheet.SetPageSetupCore(fitToWidth: fitToWidth, fitToHeight: fitToHeight);
            if (!string.IsNullOrWhiteSpace(printAreaA1)) {
                _workbook.SetPrintAreaCore(Sheet, printAreaA1!, save: false);
            }
            return this;
        }

        /// <summary>Adds a conditional color scale to the given range.</summary>
        /// <param name="rangeA1">A1-style range to format.</param>
        /// <param name="startColor">Starting color; native theme roles keep their theme slot.</param>
        /// <param name="endColor">Ending color; native theme roles keep their theme slot.</param>
        public ExcelReportBuilder ConditionalColorScale(string rangeA1, OfficeColor startColor, OfficeColor endColor) {
            Sheet.CreateConditionalColorScale(rangeA1,
                startColor.ThemeSlot != null ? startColor : startColor.Resolve(_theme.Theme),
                endColor.ThemeSlot != null ? endColor : endColor.Resolve(_theme.Theme));
            return this;
        }

        /// <summary>Adds a conditional data bar to the given range.</summary>
        /// <param name="rangeA1">A1-style range to format.</param>
        /// <param name="color">Bar color; native theme roles keep their theme slot.</param>
        public ExcelReportBuilder ConditionalDataBar(string rangeA1, OfficeColor color) {
            Sheet.CreateConditionalDataBar(rangeA1, color.ThemeSlot != null ? color : color.Resolve(_theme.Theme));
            return this;
        }

        /// <summary>Adds an icon set conditional formatting rule to the given range.</summary>
        public ExcelReportBuilder ConditionalIconSet(string rangeA1, ExcelIconSet set, bool showValue = true, bool reverse = false, double[]? percentThresholds = null, double[]? numberThresholds = null) {
            Sheet.CreateConditionalIconSet(rangeA1, set, showValue, reverse, percentThresholds, numberThresholds);
            return this;
        }

        /// <summary>Sets page orientation (Portrait/Landscape).</summary>
        public ExcelReportBuilder Orientation(OfficePageOrientation orientation) {
            Sheet.SetOrientationCore(orientation);
            return this;
        }

        /// <summary>Applies a margin preset.</summary>
        public ExcelReportBuilder Margins(ExcelMarginPreset preset) {
            Sheet.SetMarginsPresetCore(preset);
            return this;
        }

        /// <summary>Sets physical margins; omitted header/footer margins are 0.3 inches.</summary>
        public ExcelReportBuilder Margins(OfficeLength left, OfficeLength right, OfficeLength top, OfficeLength bottom, OfficeLength? header = null, OfficeLength? footer = null) {
            Sheet.SetMarginsCore(left, right, top, bottom, header, footer);
            return this;
        }

        /// <summary>Repeats header rows when printing (1-based inclusive).</summary>
        public ExcelReportBuilder RepeatHeaderRows(int firstRow, int lastRow) {
            _workbook.SetPrintTitlesCore(Sheet, firstRow, lastRow, null, null, save: false);
            return this;
        }

        /// <summary>Repeats header columns when printing (1-based inclusive).</summary>
        public ExcelReportBuilder RepeatHeaderColumns(int firstCol, int lastCol) {
            _workbook.SetPrintTitlesCore(Sheet, null, null, firstCol, lastCol, save: false);
            return this;
        }
    }
}
