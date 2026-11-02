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
            } else if (checked((uint?)(section.PageSettings.Width?.ToTwips())) > 0 && checked((uint?)(section.PageSettings.Height?.ToTwips())) > 0) {
                size = new PdfCore.PdfPageSize(
                    OfficeIMO.OfficeLength.Points(checked((uint?)(section.PageSettings.Width?.ToTwips())).GetValueOrDefault() / 20D),                    OfficeIMO.OfficeLength.Points(checked((uint?)(section.PageSettings.Height?.ToTwips())).GetValueOrDefault() / 20D));
                // The stored physical dimensions already describe the page, even when
                // the optional orientation flag is absent. Only an export override
                // should normalize these dimensions into another orientation.
                if (options?.Orientation == null && options?.DefaultOrientation == null) return size;
            } else if (section.PageSettings.PageSize.HasValue) {
                size = MapNativePageSize(section.PageSettings.PageSize.Value);
            } else if (options?.DefaultPageSize.HasValue == true) {
                size = MapNativePageSize(options.DefaultPageSize.Value);
            } else {
                size = PdfCore.PdfPageSizes.A4;
            }

            OfficePageOrientation orientation;
            if (options?.Orientation != null) {
                orientation = options.Orientation.Value;
            } else if (section.PageSettings.Orientation == OfficePageOrientation.Landscape) {
                orientation = OfficePageOrientation.Landscape;
            } else if (options?.DefaultOrientation != null) {
                orientation = options.DefaultOrientation == OfficePageOrientation.Landscape ? OfficePageOrientation.Landscape : OfficePageOrientation.Portrait;
            } else {
                orientation = OfficePageOrientation.Portrait;
            }

            return orientation == OfficePageOrientation.Landscape ? size.Landscape() : size.Portrait();
        }

        private static PdfCore.PdfPageSize MapNativePageSize(WordPageSize pageSize) =>
            WordPageSizes.GetDefinition(pageSize) is { } definition
                ? new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(checked((uint)(definition.Width.ToTwips())) / 20D),OfficeIMO.OfficeLength.Points(checked((uint)(definition.Height.ToTwips())) / 20D))
                : PdfCore.PdfPageSizes.A4;

        private static PdfCore.PdfPageMargins GetNativeMargins(WordSection section, WordPdfExportOptions? options) {
            return GetNativeMargins(section, options, GetNativeHeaderFooterMarginExpansion(section, options));
        }

        private static PdfCore.PdfPageMargins GetNativeMargins(WordSection section, WordPdfExportOptions? options, (double Header, double Footer) headerFooterMarginExpansion) {
            if (options?.Margins != null) {
                return options.Margins.Value;
            }

            double left = checked((uint)(section.Margins.Left.GetValueOrDefault().ToTwips())) / 20D;
            double top = (checked((int?)(section.Margins.Top?.ToTwips())) ?? 0) / 20D + headerFooterMarginExpansion.Header;
            double right = checked((uint)(section.Margins.Right.GetValueOrDefault().ToTwips())) / 20D;
            double bottom = (checked((int?)(section.Margins.Bottom?.ToTwips())) ?? 0) / 20D + headerFooterMarginExpansion.Footer;
            double gutter = checked((uint)(section.Margins.Gutter.GetValueOrDefault().ToTwips())) / 20D;
            if (section._document.Settings.GutterAtTop) top += gutter;
            else if (section.RtlGutter) right += gutter;
            else left += gutter;
            return new PdfCore.PdfPageMargins(OfficeIMO.OfficeLength.Points(left),OfficeIMO.OfficeLength.Points(top),OfficeIMO.OfficeLength.Points(right),OfficeIMO.OfficeLength.Points(bottom));
        }
    }
}
