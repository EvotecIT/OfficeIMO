namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Print and conditional formatting conveniences for SheetComposer.
    /// </summary>
    public sealed partial class SheetComposer {
        /// <summary>
        /// Applies sensible print defaults: gridlines off, fit to width, and optional print area.
        /// </summary>
        public SheetComposer PrintDefaults(bool showGridlines = false, uint fitToWidth = 1, uint fitToHeight = 0, string? printAreaA1 = null) {
            Sheet.SetGridlinesVisible(showGridlines);
            Sheet.SetPageSetup(fitToWidth: fitToWidth, fitToHeight: fitToHeight);
            if (!string.IsNullOrWhiteSpace(printAreaA1)) {
                try { _workbook.SetPrintArea(Sheet, printAreaA1!, save: false); } catch { }
            }
            return this;
        }

        /// <summary>Adds a conditional color scale to the given range.</summary>
        public SheetComposer ConditionalColorScale(string rangeA1, string startHex, string endHex) {
            try { Sheet.AddConditionalColorScale(rangeA1, startHex, endHex); } catch { }
            return this;
        }

        /// <summary>Adds a conditional data bar to the given range.</summary>
        public SheetComposer ConditionalDataBar(string rangeA1, string colorHex) {
            try { Sheet.AddConditionalDataBar(rangeA1, colorHex); } catch { }
            return this;
        }

        /// <summary>Adds an icon set conditional formatting rule to the given range.</summary>
        public SheetComposer ConditionalIconSet(string rangeA1, DocumentFormat.OpenXml.Spreadsheet.IconSetValues set, bool showValue = true, bool reverse = false, double[]? percentThresholds = null, double[]? numberThresholds = null) {
            try { Sheet.AddConditionalIconSet(rangeA1, set, showValue, reverse, percentThresholds, numberThresholds); } catch { }
            return this;
        }

        /// <summary>Sets page orientation (Portrait/Landscape).</summary>
        public SheetComposer Orientation(ExcelPageOrientation orientation) {
            try { Sheet.SetOrientation(orientation); } catch { }
            return this;
        }

        /// <summary>Applies a margin preset.</summary>
        public SheetComposer Margins(ExcelMarginPreset preset) {
            try { Sheet.SetMarginsPreset(preset); } catch { }
            return this;
        }

        /// <summary>Sets explicit margins in inches.</summary>
        public SheetComposer Margins(double left, double right, double top, double bottom, double header = 0.3, double footer = 0.3) {
            try { Sheet.SetMargins(left, right, top, bottom, header, footer); } catch { }
            return this;
        }

        /// <summary>Repeats header rows when printing (1-based inclusive).</summary>
        public SheetComposer RepeatHeaderRows(int firstRow, int lastRow) {
            try { _workbook.SetPrintTitles(Sheet, firstRow, lastRow, null, null, save: false); } catch { }
            return this;
        }

        /// <summary>Repeats header columns when printing (1-based inclusive).</summary>
        public SheetComposer RepeatHeaderColumns(int firstCol, int lastCol) {
            try { _workbook.SetPrintTitles(Sheet, null, null, firstCol, lastCol, save: false); } catch { }
            return this;
        }
    }
}
