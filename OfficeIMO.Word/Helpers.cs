using System.Diagnostics;

namespace OfficeIMO.Word {
    public static partial class Helpers {
        /// <summary>
        /// Converts Color to Hex Color
        /// </summary>
        /// <param name="c"></param>
        /// <returns></returns>
        public static string ToHexColor(this SixLabors.ImageSharp.Color c) {
            return c.ToHex().Remove(6);
        }

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
