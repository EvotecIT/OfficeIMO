namespace OfficeIMO.Excel {
    /// <summary>
    /// Provides helper methods used by <c>OfficeIMO.Excel</c> components.
    /// </summary>
    public static partial class Helpers {

        /// <summary>
        /// Converts a <see cref="OfficeIMO.Drawing.OfficeColor"/> to a hexadecimal
        /// color string.
        /// </summary>
        /// <param name="c">Color to convert.</param>
        /// <returns>Hexadecimal representation of the color.</returns>
        public static string ToHexColor(this OfficeIMO.Drawing.OfficeColor c) {
            return c.ToHex().Remove(6);
        }

    }
}
