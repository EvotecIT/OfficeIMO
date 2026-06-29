using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.LegacyDoc.Write {
    internal static partial class LegacyDocWriter {
        private const int DefaultPageWidthTwips = 12240;
        private const int DefaultPageHeightTwips = 15840;
        private const int DefaultPageMarginTwips = 1440;
        private const int DefaultHeaderFooterMarginTwips = 720;

        private static void ThrowIfUnsupportedSectionProperties(SectionProperties sectionProperties) {
            foreach (OpenXmlElement property in sectionProperties.ChildElements) {
                switch (property) {
                    case PageSize pageSize:
                        ThrowIfUnsupportedPageSize(pageSize);
                        break;
                    case PageMargin pageMargin:
                        ThrowIfUnsupportedPageMargin(pageMargin);
                        break;
                    default:
                        throw new NotSupportedException($"Native DOC saving currently supports only the default final section page setup. Unsupported section property: {property.LocalName}.");
                }
            }
        }

        private static void ThrowIfUnsupportedPageSize(PageSize pageSize) {
            PageOrientationValues orientation = pageSize.Orient?.Value ?? PageOrientationValues.Portrait;
            if (orientation != PageOrientationValues.Portrait
                || !IsTwipValue(pageSize.Width, DefaultPageWidthTwips)
                || !IsTwipValue(pageSize.Height, DefaultPageHeightTwips)) {
                throw new NotSupportedException("Native DOC saving currently supports only the default Letter portrait section page setup.");
            }
        }

        private static void ThrowIfUnsupportedPageMargin(PageMargin pageMargin) {
            if (!IsTwipValue(pageMargin.Top, DefaultPageMarginTwips)
                || !IsTwipValue(pageMargin.Right, DefaultPageMarginTwips)
                || !IsTwipValue(pageMargin.Bottom, DefaultPageMarginTwips)
                || !IsTwipValue(pageMargin.Left, DefaultPageMarginTwips)
                || !IsTwipValue(pageMargin.Header, DefaultHeaderFooterMarginTwips)
                || !IsTwipValue(pageMargin.Footer, DefaultHeaderFooterMarginTwips)
                || !IsTwipValue(pageMargin.Gutter, 0)) {
                throw new NotSupportedException("Native DOC saving currently supports only the default final section margins.");
            }
        }

        private static bool IsTwipValue(OpenXmlSimpleType? value, int expected) {
            if (value == null) {
                return true;
            }

            return int.TryParse(value.InnerText, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out int actual)
                && actual == expected;
        }
    }
}
