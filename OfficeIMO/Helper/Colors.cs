using System;

namespace OfficeIMO.Helper {
    public static class Colors {
        public static string ToHexColor(this System.Drawing.Color c) {
            return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }
        public static string ToRgbColor(this System.Drawing.Color c) => $"RGB({c.R}, {c.G}, {c.B})";
    }
}
