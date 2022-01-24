namespace OfficeIMO {
    public static partial class WordHelpers {

        /// <summary>
        /// Converts Color to Hex Color
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public static string ToHexColor(this System.Drawing.Color c) {
            return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }
        /// <summary>
        /// Converts Color to RGB Color
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public static string ToRgbColor(this System.Drawing.Color c) => $"RGB({c.R}, {c.G}, {c.B})";
    }
}
