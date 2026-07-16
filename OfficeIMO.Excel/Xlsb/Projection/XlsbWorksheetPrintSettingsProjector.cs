using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares standard XLSB worksheet print options and margins.</summary>
    internal static class XlsbWorksheetPrintSettingsProjector {
        internal static void Apply(
            ExcelSheet sheet,
            XlsbPrintOptions? options,
            XlsbPageMargins? margins,
            XlsbPageSetup? pageSetup,
            XlsbHeaderFooter? headerFooter) {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            if (options != null) worksheet.Append(Create(options));
            if (margins != null) worksheet.Append(Create(margins));
            if (pageSetup != null) worksheet.Append(Create(pageSetup));
            if (headerFooter != null) worksheet.Append(Create(headerFooter));
        }

        internal static void ValidateUnchanged(
            ExcelSheet sheet,
            XlsbPrintOptions? expectedOptions,
            XlsbPageMargins? expectedMargins,
            XlsbPageSetup? expectedPageSetup,
            XlsbHeaderFooter? expectedHeaderFooter) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet
                ?? throw new InvalidDataException($"Worksheet '{sheet.Name}' has no worksheet root.");
            ValidateSingle(worksheet.Elements<PrintOptions>().ToArray(), expectedOptions == null ? null : Create(expectedOptions), sheet, "print options");
            ValidateSingle(worksheet.Elements<PageMargins>().ToArray(), expectedMargins == null ? null : Create(expectedMargins), sheet, "page margins");
            ValidateSingle(worksheet.Elements<PageSetup>().ToArray(), expectedPageSetup == null ? null : Create(expectedPageSetup), sheet, "page setup");
            ValidateSingle(worksheet.Elements<HeaderFooter>().ToArray(), expectedHeaderFooter == null ? null : Create(expectedHeaderFooter), sheet, "headers and footers");
        }

        private static PrintOptions Create(XlsbPrintOptions source) => new PrintOptions {
            HorizontalCentered = source.HorizontalCentered,
            VerticalCentered = source.VerticalCentered,
            Headings = source.Headings,
            GridLines = source.GridLines,
            GridLinesSet = true
        };

        private static PageMargins Create(XlsbPageMargins source) => new PageMargins {
            Left = source.Left,
            Right = source.Right,
            Top = source.Top,
            Bottom = source.Bottom,
            Header = source.Header,
            Footer = source.Footer
        };

        private static PageSetup Create(XlsbPageSetup source) {
            var result = new PageSetup {
                FitToWidth = source.FitToWidth,
                FitToHeight = source.FitToHeight,
                PageOrder = source.OverThenDown ? PageOrderValues.OverThenDown : PageOrderValues.DownThenOver,
                Orientation = source.UseDefaultOrientation
                    ? OrientationValues.Default
                    : source.Landscape ? OrientationValues.Landscape : OrientationValues.Portrait,
                UsePrinterDefaults = source.UseDefaultOrientation,
                BlackAndWhite = source.BlackAndWhite,
                Draft = source.Draft,
                CellComments = !source.PrintCellComments
                    ? CellCommentsValues.None
                    : source.CommentsAtEnd ? CellCommentsValues.AtEnd : CellCommentsValues.AsDisplayed,
                UseFirstPageNumber = source.UseFirstPageNumber,
                Errors = source.Errors switch {
                    XlsbPrintErrorMode.Blank => PrintErrorValues.Blank,
                    XlsbPrintErrorMode.Dash => PrintErrorValues.Dash,
                    XlsbPrintErrorMode.NotAvailable => PrintErrorValues.NA,
                    _ => PrintErrorValues.Displayed
                }
            };
            if (source.PaperSize != 0U) result.PaperSize = source.PaperSize;
            if (source.Scale != 0U) result.Scale = source.Scale;
            if (source.HorizontalDpi != 0U) result.HorizontalDpi = source.HorizontalDpi;
            if (source.VerticalDpi != 0U) result.VerticalDpi = source.VerticalDpi;
            if (source.Copies != 0U) result.Copies = source.Copies;
            if (source.UseFirstPageNumber && source.FirstPageNumber >= 0) {
                result.FirstPageNumber = checked((uint)source.FirstPageNumber);
            }
            return result;
        }

        private static HeaderFooter Create(XlsbHeaderFooter source) {
            var result = new HeaderFooter {
                DifferentOddEven = source.DifferentOddEven,
                DifferentFirst = source.DifferentFirst,
                ScaleWithDoc = source.ScaleWithDocument,
                AlignWithMargins = source.AlignWithMargins
            };
            if (source.OddHeader != null) result.Append(new OddHeader(source.OddHeader));
            if (source.OddFooter != null) result.Append(new OddFooter(source.OddFooter));
            if (source.EvenHeader != null) result.Append(new EvenHeader(source.EvenHeader));
            if (source.EvenFooter != null) result.Append(new EvenFooter(source.EvenFooter));
            if (source.FirstHeader != null) result.Append(new FirstHeader(source.FirstHeader));
            if (source.FirstFooter != null) result.Append(new FirstFooter(source.FirstFooter));
            return result;
        }

        private static void ValidateSingle<TElement>(TElement[] actual, TElement? expected, ExcelSheet sheet, string detail)
            where TElement : DocumentFormat.OpenXml.OpenXmlElement {
            if (actual.Length > 1
                || (expected == null && actual.Length != 0)
                || (expected != null
                    && (actual.Length != 1
                        || !string.Equals(actual[0].OuterXml, expected.OuterXml, StringComparison.Ordinal)))) {
                throw new NotSupportedException($"Native XLSB rewriting preserves but cannot modify {detail} on worksheet '{sheet.Name}'. Save as .xlsx to retain that change.");
            }
        }
    }
}
