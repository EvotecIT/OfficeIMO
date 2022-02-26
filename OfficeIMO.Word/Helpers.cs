using System.Diagnostics;

namespace OfficeIMO.Word {
    public static partial class Helpers {

        /// <summary>
        /// Converts Color to Hex Color
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public static string ToHexColor(this System.Drawing.Color c) {
            return c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
            //return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }
        /// <summary>
        /// Converts Color to RGB Color
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public static string ToRgbColor(this System.Drawing.Color c) => $"RGB({c.R}, {c.G}, {c.B})";



        /// <summary>
        /// Opens up any file using assigned Application
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="open"></param>
        public static void Open(string filePath, bool open) {
            if (open) {
                ProcessStartInfo startInfo = new ProcessStartInfo(filePath) {
                    UseShellExecute = true
                };
                Process.Start(startInfo);
            }
        }
    }
}
