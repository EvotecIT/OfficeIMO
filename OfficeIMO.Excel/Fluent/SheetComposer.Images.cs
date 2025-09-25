namespace OfficeIMO.Excel.Fluent {
    /// <summary>
    /// Image helpers for SheetComposer that forward to the underlying sheet and keep chaining.
    /// </summary>
    public sealed partial class SheetComposer {
        /// <summary>
        /// Inserts an image anchored to a specific cell.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="bytes">Image bytes.</param>
        /// <param name="contentType">Image content type (e.g., image/png).</param>
        /// <param name="widthPixels">Width in pixels.</param>
        /// <param name="heightPixels">Height in pixels.</param>
        /// <param name="offsetXPixels">Optional X offset in pixels.</param>
        /// <param name="offsetYPixels">Optional Y offset in pixels.</param>
        public SheetComposer ImageAt(int row, int column, byte[] bytes, string contentType = "image/png", int widthPixels = 96, int heightPixels = 32, int offsetXPixels = 0, int offsetYPixels = 0) {
            Sheet.AddImageAt(row, column, bytes, contentType, widthPixels, heightPixels, offsetXPixels, offsetYPixels);
            return this;
        }

        /// <summary>
        /// Downloads an image from URL and inserts it anchored at a specific cell.
        /// </summary>
        /// <param name="row">1-based row index.</param>
        /// <param name="column">1-based column index.</param>
        /// <param name="url">Direct URL to an image.</param>
        /// <param name="widthPixels">Width in pixels.</param>
        /// <param name="heightPixels">Height in pixels.</param>
        /// <param name="offsetXPixels">Optional X offset in pixels.</param>
        /// <param name="offsetYPixels">Optional Y offset in pixels.</param>
        public SheetComposer ImageFromUrlAt(int row, int column, string url, int widthPixels = 96, int heightPixels = 32, int offsetXPixels = 0, int offsetYPixels = 0) {
            Sheet.AddImageFromUrlAt(row, column, url, widthPixels, heightPixels, offsetXPixels, offsetYPixels);
            return this;
        }
    }
}
