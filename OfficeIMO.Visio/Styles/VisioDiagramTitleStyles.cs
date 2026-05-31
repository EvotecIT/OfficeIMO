using System;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Visio {
    internal static class VisioDiagramTitleStyles {
        public static VisioTextStyle Create(VisioStyleTheme theme, double minimumSize = 20D) {
            VisioTextStyle style = theme.Emphasis.TextStyle?.Clone() ?? new VisioTextStyle();
            style.FontFamily = string.IsNullOrWhiteSpace(style.FontFamily) ? "Aptos Display" : style.FontFamily;
            style.Size = Math.Max(style.Size ?? 0D, minimumSize);
            style.Bold = true;
            style.Color = Color.FromRgb(32, 55, 75);
            style.HorizontalAlignment = VisioTextHorizontalAlignment.Center;
            style.VerticalAlignment = VisioTextVerticalAlignment.Middle;
            return style;
        }
    }
}
