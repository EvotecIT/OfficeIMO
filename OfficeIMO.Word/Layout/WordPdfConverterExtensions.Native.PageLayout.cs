using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static bool ShouldMirrorNativeMargins(WordSection section, WordToPdfOptions? options) =>
            options?.Margins == null && section._document.Settings.MirrorMargins &&
            !section._document.Settings.GutterAtTop;

        private static PdfCore.PageSize GetNativePageSize(WordSection section, WordToPdfOptions? options) {
            PdfCore.PageSize size;
            if (options?.PageSize != null) {
                size = options.PageSize.Value;
                if (options.Orientation == null) {
                    return size;
                }
            } else if (section.PageSettings.Width > 0 && section.PageSettings.Height > 0) {
                size = new PdfCore.PageSize(
                    section.PageSettings.Width.GetValueOrDefault() / 20D,
                    section.PageSettings.Height.GetValueOrDefault() / 20D);
                // The stored physical dimensions already describe the page, even when
                // the optional orientation flag is absent. Only an export override
                // should normalize these dimensions into another orientation.
                if (options?.Orientation == null && options?.DefaultOrientation == null) return size;
            } else if (section.PageSettings.PageSize.HasValue) {
                size = MapNativePageSize(section.PageSettings.PageSize.Value);
            } else if (options?.DefaultPageSize.HasValue == true) {
                size = MapNativePageSize(options.DefaultPageSize.Value);
            } else {
                size = PdfCore.PageSizes.A4;
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

        private static PdfCore.PageSize MapNativePageSize(WordPageSize pageSize) =>
            WordPageSizes.GetDefinition(pageSize) is { } definition
                ? new PdfCore.PageSize(definition.WidthTwips / 20D, definition.HeightTwips / 20D)
                : PdfCore.PageSizes.A4;

        private static PdfCore.PageMargins GetNativeMargins(WordSection section, WordToPdfOptions? options) {
            return GetNativeMargins(section, options, GetNativeHeaderFooterMarginExpansion(section, options));
        }

        private static PdfCore.PageMargins GetNativeMargins(WordSection section, WordToPdfOptions? options, (double Header, double Footer) headerFooterMarginExpansion) {
            if (options?.Margins != null) {
                return options.Margins.Value;
            }

            double left = section.Margins.Left / 20D;
            double top = (section.Margins.Top ?? 0) / 20D + headerFooterMarginExpansion.Header;
            double right = section.Margins.Right / 20D;
            double bottom = (section.Margins.Bottom ?? 0) / 20D + headerFooterMarginExpansion.Footer;
            double gutter = section.Margins.Gutter / 20D;
            if (section._document.Settings.GutterAtTop) top += gutter;
            else if (section.RtlGutter) right += gutter;
            else left += gutter;
            return new PdfCore.PageMargins(left, top, right, bottom);
        }
    }
}
