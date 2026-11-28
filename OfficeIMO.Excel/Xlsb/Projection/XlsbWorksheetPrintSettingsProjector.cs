using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Projection {
    /// <summary>Projects and compares XLSB print settings in SpreadsheetML worksheet markup.</summary>
    internal static class XlsbWorksheetPrintSettingsProjector {
        internal const string Options = "printOptions";
        internal const string Margins = "pageMargins";
        internal const string Setup = "pageSetup";
        internal const string Headers = "headerFooter";

        internal static void Apply(ExcelSheet sheet, XlsbPrintOptions? options, XlsbPageMargins? margins,
            XlsbPageSetup? pageSetup, XlsbHeaderFooter? headerFooter) =>
            Apply(XlsbWorksheetMarkup.Root(sheet), options, margins, pageSetup, headerFooter);

        internal static void Apply(OfficeXmlNode worksheet, XlsbPrintOptions? options, XlsbPageMargins? margins,
            XlsbPageSetup? pageSetup, XlsbHeaderFooter? headerFooter) {
            if (options != null) worksheet.Append(Create(worksheet, options));
            if (margins != null) worksheet.Append(Create(worksheet, margins));
            if (pageSetup != null) worksheet.Append(Create(worksheet, pageSetup));
            if (headerFooter != null) worksheet.Append(Create(worksheet, headerFooter));
        }

        internal static void ValidateUnchanged(ExcelSheet sheet, XlsbPrintOptions? options, XlsbPageMargins? margins,
            XlsbPageSetup? pageSetup, XlsbHeaderFooter? headerFooter) {
            OfficeXmlNode worksheet = XlsbWorksheetMarkup.Root(sheet);
            Validate(Options, options == null ? default : Create(worksheet, options), "print options");
            Validate(Margins, margins == null ? default : Create(worksheet, margins), "page margins");
            Validate(Setup, pageSetup == null ? default : Create(worksheet, pageSetup), "page setup");
            Validate(Headers, headerFooter == null ? default : Create(worksheet, headerFooter), "headers and footers");

            void Validate(string elementName, OfficeXmlNode expected, string detail) {
                if (!XlsbWorksheetMarkup.MatchesChild(worksheet, elementName, expected)) {
                    throw new NotSupportedException($"Native XLSB rewriting preserves but cannot modify {detail} on worksheet '{sheet.Name}'. Save as .xlsx to retain that change.");
                }
            }
        }

        private static OfficeXmlNode Create(OfficeXmlNode worksheet, XlsbPrintOptions source) {
            OfficeXmlNode result = XlsbWorksheetMarkup.Create(worksheet, Options);
            result.SetBooleanAttribute("horizontalCentered", source.HorizontalCentered);
            result.SetBooleanAttribute("verticalCentered", source.VerticalCentered);
            result.SetBooleanAttribute("headings", source.Headings);
            result.SetBooleanAttribute("gridLines", source.GridLines);
            result.SetBooleanAttribute("gridLinesSet", true);
            return result;
        }

        private static OfficeXmlNode Create(OfficeXmlNode worksheet, XlsbPageMargins source) {
            OfficeXmlNode result = XlsbWorksheetMarkup.Create(worksheet, Margins);
            result.SetDoubleAttribute("left", source.Left);
            result.SetDoubleAttribute("right", source.Right);
            result.SetDoubleAttribute("top", source.Top);
            result.SetDoubleAttribute("bottom", source.Bottom);
            result.SetDoubleAttribute("header", source.Header);
            result.SetDoubleAttribute("footer", source.Footer);
            return result;
        }

        private static OfficeXmlNode Create(OfficeXmlNode worksheet, XlsbPageSetup source) {
            OfficeXmlNode result = XlsbWorksheetMarkup.Create(worksheet, Setup);
            result.SetUInt32Attribute("fitToWidth", source.FitToWidth);
            result.SetUInt32Attribute("fitToHeight", source.FitToHeight);
            result.SetAttribute("pageOrder", source.OverThenDown ? "overThenDown" : "downThenOver");
            result.SetAttribute("orientation", source.UseDefaultOrientation ? "default" : source.Landscape ? "landscape" : "portrait");
            result.SetBooleanAttribute("usePrinterDefaults", source.UseDefaultOrientation);
            result.SetBooleanAttribute("blackAndWhite", source.BlackAndWhite);
            result.SetBooleanAttribute("draft", source.Draft);
            result.SetAttribute("cellComments", !source.PrintCellComments ? "none" : source.CommentsAtEnd ? "atEnd" : "asDisplayed");
            result.SetBooleanAttribute("useFirstPageNumber", source.UseFirstPageNumber);
            result.SetAttribute("errors", source.Errors switch {
                XlsbPrintErrorMode.Blank => "blank",
                XlsbPrintErrorMode.Dash => "dash",
                XlsbPrintErrorMode.NotAvailable => "NA",
                _ => "displayed"
            });
            if (source.PaperSize != 0U) result.SetUInt32Attribute("paperSize", source.PaperSize);
            if (source.Scale != 0U) result.SetUInt32Attribute("scale", source.Scale);
            if (source.HorizontalDpi != 0U) result.SetUInt32Attribute("horizontalDpi", source.HorizontalDpi);
            if (source.VerticalDpi != 0U) result.SetUInt32Attribute("verticalDpi", source.VerticalDpi);
            if (source.Copies != 0U) result.SetUInt32Attribute("copies", source.Copies);
            if (source.UseFirstPageNumber && source.FirstPageNumber >= 0) {
                result.SetUInt32Attribute("firstPageNumber", checked((uint)source.FirstPageNumber));
            }
            return result;
        }

        private static OfficeXmlNode Create(OfficeXmlNode worksheet, XlsbHeaderFooter source) {
            OfficeXmlNode result = XlsbWorksheetMarkup.Create(worksheet, Headers);
            result.SetBooleanAttribute("differentOddEven", source.DifferentOddEven);
            result.SetBooleanAttribute("differentFirst", source.DifferentFirst);
            result.SetBooleanAttribute("scaleWithDoc", source.ScaleWithDocument);
            result.SetBooleanAttribute("alignWithMargins", source.AlignWithMargins);
            AppendText(result, "oddHeader", source.OddHeader);
            AppendText(result, "oddFooter", source.OddFooter);
            AppendText(result, "evenHeader", source.EvenHeader);
            AppendText(result, "evenFooter", source.EvenFooter);
            AppendText(result, "firstHeader", source.FirstHeader);
            AppendText(result, "firstFooter", source.FirstFooter);
            return result;
        }

        private static void AppendText(OfficeXmlNode parent, string localName, string? text) {
            if (text == null) return;
            OfficeXmlNode child = XlsbWorksheetMarkup.Create(parent, localName);
            child.SetInnerText(text);
            parent.Append(child);
        }
    }
}
