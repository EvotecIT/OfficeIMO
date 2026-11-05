using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool ShouldMirrorNativeMargins(WordSection section, WordPdfExportOptions? options) =>
            options?.Margins == null && section._document.Settings.MirrorMargins &&
            !section._document.Settings.GutterAtTop;

        private static PdfCore.PdfPageSize GetNativePageSize(WordSection section, WordPdfExportOptions? options) {
            PdfCore.PdfPageSize size;
            if (options?.PageSize != null) {
                size = options.PageSize.Value;
                if (options.Orientation == null) {
                    return size;
                }
            } else if (checked((uint?)(section.PageSetup.Size.Width?.ToTwips())) > 0 && checked((uint?)(section.PageSetup.Size.Height?.ToTwips())) > 0) {
                size = new PdfCore.PdfPageSize(
                    OfficeIMO.OfficeLength.Points(checked((uint?)(section.PageSetup.Size.Width?.ToTwips())).GetValueOrDefault() / 20D),                    OfficeIMO.OfficeLength.Points(checked((uint?)(section.PageSetup.Size.Height?.ToTwips())).GetValueOrDefault() / 20D));
                // The stored physical dimensions already describe the page, even when
                // the optional orientation flag is absent. Only an export override
                // should normalize these dimensions into another orientation.
                if (options?.Orientation == null && options?.FallbackOrientation == null) return size;
            } else if (WordPageSizes.GetDefinition(section.PageSetup.Size.PageSize.GetValueOrDefault()) is { } definition) {
                size = new PdfCore.PdfPageSize(definition.Width, definition.Height);
            } else if (options?.FallbackPageSize.HasValue == true) {
                size = options.FallbackPageSize.Value;
                // Physical fallback dimensions already carry their orientation. An absent
                // source flag must not manufacture a portrait override.
                if (options.Orientation == null && options.FallbackOrientation == null &&
                    section._sectionProperties.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.PageSize>()?.Orient == null) return size;
            } else {
                size = PdfCore.PdfPageSizes.A4;
            }

            OfficePageOrientation orientation;
            if (options?.Orientation != null) {
                orientation = options.Orientation.Value;
            } else if (section.PageSetup.Size.Orientation == OfficePageOrientation.Landscape) {
                orientation = OfficePageOrientation.Landscape;
            } else if (options?.FallbackOrientation != null) {
                orientation = options.FallbackOrientation == OfficePageOrientation.Landscape ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;
            } else {
                orientation = OfficePageOrientation.Portrait;
            }

            return orientation == OfficePageOrientation.Landscape ? size.Landscape() : size.Portrait();
        }

        private static PdfCore.PdfPageMargins GetNativeMargins(WordSection section, WordPdfExportOptions? options) {
            return GetNativeMargins(section, options, GetNativeHeaderFooterMarginExpansion(section, options));
        }

        private static PdfCore.PdfPageMargins GetNativeMargins(WordSection section, WordPdfExportOptions? options, (double Header, double Footer) headerFooterMarginExpansion) {
            if (options?.Margins != null) {
                return options.Margins.Value;
            }

            double left = checked((uint)(section.PageSetup.Margins.Left.GetValueOrDefault().ToTwips())) / 20D;
            double top = (checked((int?)(section.PageSetup.Margins.Top?.ToTwips())) ?? 0) / 20D + headerFooterMarginExpansion.Header;
            double right = checked((uint)(section.PageSetup.Margins.Right.GetValueOrDefault().ToTwips())) / 20D;
            double bottom = (checked((int?)(section.PageSetup.Margins.Bottom?.ToTwips())) ?? 0) / 20D + headerFooterMarginExpansion.Footer;
            double gutter = checked((uint)(section.PageSetup.Margins.Gutter.GetValueOrDefault().ToTwips())) / 20D;
            if (section._document.Settings.GutterAtTop) top += gutter;
            else if (section.RtlGutter) right += gutter;
            else left += gutter;
            return new PdfCore.PdfPageMargins(OfficeIMO.OfficeLength.Points(left),OfficeIMO.OfficeLength.Points(top),OfficeIMO.OfficeLength.Points(right),OfficeIMO.OfficeLength.Points(bottom));
        }
    }
}
