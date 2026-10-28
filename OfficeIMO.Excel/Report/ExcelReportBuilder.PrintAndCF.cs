namespace OfficeIMO.Excel {
    // Print and conditional formatting conveniences for ExcelReportBuilder.
    public sealed partial class ExcelReportBuilder {
        /// <summary>
        /// Applies sensible print defaults: gridlines off, fit to width, and optional print area.
        /// </summary>
        public ExcelReportBuilder PrintDefaults(bool showGridlines = false, uint fitToWidth = 1, uint fitToHeight = 0, string? printAreaA1 = null) {
            Sheet.SetGridlinesVisible(showGridlines);
            Sheet.SetPageSetup(fitToWidth: fitToWidth, fitToHeight: fitToHeight);
            if (!string.IsNullOrWhiteSpace(printAreaA1)) {
                _workbook.SetPrintArea(Sheet, printAreaA1!, save: false);
            }
            return this;
        }

        /// <summary>Adds a conditional color scale to the given range.</summary>
        /// <param name="rangeA1">A1-style range to format.</param>
        /// <param name="startColor">Starting color; native theme roles keep their theme slot.</param>
        /// <param name="endColor">Ending color; native theme roles keep their theme slot.</param>
        public ExcelReportBuilder ConditionalColorScale(string rangeA1, OfficeColor startColor, OfficeColor endColor) {
            Sheet.AddConditionalColorScale(rangeA1,
                startColor.ThemeSlot != null ? startColor : startColor.Resolve(_theme.Theme),
                endColor.ThemeSlot != null ? endColor : endColor.Resolve(_theme.Theme));
            return this;
        }

        /// <summary>Adds a conditional data bar to the given range.</summary>
        /// <param name="rangeA1">A1-style range to format.</param>
        /// <param name="color">Bar color; native theme roles keep their theme slot.</param>
        public ExcelReportBuilder ConditionalDataBar(string rangeA1, OfficeColor color) {
            Sheet.AddConditionalDataBar(rangeA1, color.ThemeSlot != null ? color : color.Resolve(_theme.Theme));
            return this;
        }

        /// <summary>Adds an icon set conditional formatting rule to the given range.</summary>
        public ExcelReportBuilder ConditionalIconSet(string rangeA1, ExcelIconSet set, bool showValue = true, bool reverse = false, double[]? percentThresholds = null, double[]? numberThresholds = null) {
            Sheet.AddConditionalIconSet(rangeA1, set, showValue, reverse, percentThresholds, numberThresholds);
            return this;
        }

        /// <summary>Sets page orientation (Portrait/Landscape).</summary>
        public ExcelReportBuilder Orientation(OfficePageOrientation orientation) {
            Sheet.SetOrientation(orientation);
            return this;
        }

        /// <summary>Applies a margin preset.</summary>
        public ExcelReportBuilder Margins(ExcelMarginPreset preset) {
            Sheet.SetMarginsPreset(preset);
            return this;
        }

        /// <summary>Sets explicit margins in inches.</summary>
        public ExcelReportBuilder Margins(double left, double right, double top, double bottom, double header = 0.3, double footer = 0.3) {
            Sheet.SetMargins(left, right, top, bottom, header, footer);
            return this;
        }

        /// <summary>Repeats header rows when printing (1-based inclusive).</summary>
        public ExcelReportBuilder RepeatHeaderRows(int firstRow, int lastRow) {
            _workbook.SetPrintTitles(Sheet, firstRow, lastRow, null, null, save: false);
            return this;
        }

        /// <summary>Repeats header columns when printing (1-based inclusive).</summary>
        public ExcelReportBuilder RepeatHeaderColumns(int firstCol, int lastCol) {
            _workbook.SetPrintTitles(Sheet, null, null, firstCol, lastCol, save: false);
            return this;
        }
    }
}
