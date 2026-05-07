using OfficeIMO.Drawing;

namespace OfficeIMO.Excel {
    internal sealed class ExcelTextMeasurer {
        private readonly OfficeTextMeasurer _textMeasurer;

        private ExcelTextMeasurer(OfficeFontInfo fallbackFontInfo) {
            _textMeasurer = OfficeTextMeasurer.Create(fallbackFontInfo);
            DefaultStyle = new Style(_textMeasurer.DefaultStyle);
        }

        internal Style DefaultStyle { get; }

        internal float DefaultFontSize => (float)_textMeasurer.DefaultFontSize;

        internal OfficeFontInfo FallbackFontInfo => _textMeasurer.FallbackFontInfo;

        internal static ExcelTextMeasurer Create(OfficeFontInfo? workbookDefaultFontInfo) =>
            new ExcelTextMeasurer(workbookDefaultFontInfo ?? OfficeFontInfo.Default);

        internal Style CreateDefaultStyle(float dpi)
            => new Style(_textMeasurer.CreateDefaultStyle(dpi));

        internal Style CreateStyle(OfficeFontInfo fontInfo)
            => new Style(_textMeasurer.CreateStyle(fontInfo));

        internal Style CreateStyle(OfficeFontInfo fontInfo, float dpi)
            => new Style(_textMeasurer.CreateStyle(fontInfo, dpi));

        internal float MeasureWidthOrDefault(string text, Style style, float fallback) {
            float measured = (float)_textMeasurer.MeasureWidthOrDefault(text, style.DrawingStyle, fallback);
            return measured > 0.0001f ? measured : fallback;
        }

        internal float MeasureHeightOrDefault(string text, Style style, float fallback) {
            float measured = (float)_textMeasurer.MeasureLineHeightOrDefault(text, style.DrawingStyle, fallback);
            return measured > 0.0001f ? measured : fallback;
        }

        internal readonly struct Style {
            internal Style(OfficeTextMeasurementStyle drawingStyle) {
                DrawingStyle = drawingStyle;
            }

            internal OfficeTextMeasurementStyle DrawingStyle { get; }

            internal OfficeFontInfo FontInfo => DrawingStyle.FontInfo;

            internal float Dpi => (float)DrawingStyle.Dpi;

            internal float FontSizePixels => (float)DrawingStyle.FontSizePixels;

            internal float SpaceWidth => (float)DrawingStyle.SpaceWidthPixels;

            internal float MaximumDigitWidth => (float)DrawingStyle.MaximumDigitWidthPixels;
        }
    }
}
