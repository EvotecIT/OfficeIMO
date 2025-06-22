using System.Diagnostics;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Provides helper methods used by <c>OfficeIMO.Excel</c> components.
    /// </summary>
    public static partial class Helpers {

        /// <summary>
        /// Converts a <see cref="SixLabors.ImageSharp.Color"/> to a hexadecimal
        /// color string.
        /// </summary>
        /// <param name="c">Color to convert.</param>
        /// <returns>Hexadecimal representation of the color.</returns>
        public static string ToHexColor(this SixLabors.ImageSharp.Color c) {
            return c.ToHex().Remove(6);
        }

        /// <summary>
        /// Opens the specified file using the default application.
        /// </summary>
        /// <param name="filePath">Path to the file to open.</param>
        /// <param name="open">When <c>true</c>, the file is opened.</param>
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
