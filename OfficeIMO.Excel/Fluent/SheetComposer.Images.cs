namespace OfficeIMO.Excel.Fluent
{
    /// <summary>
    /// Image helpers for SheetComposer that forward to the underlying sheet and keep chaining.
    /// </summary>
    public sealed partial class SheetComposer
    {
        public SheetComposer ImageAt(int row, int column, byte[] bytes, string contentType = "image/png", int widthPixels = 96, int heightPixels = 32, int offsetXPixels = 0, int offsetYPixels = 0)
        {
            Sheet.AddImageAt(row, column, bytes, contentType, widthPixels, heightPixels, offsetXPixels, offsetYPixels);
            return this;
        }

        public SheetComposer ImageFromUrlAt(int row, int column, string url, int widthPixels = 96, int heightPixels = 32, int offsetXPixels = 0, int offsetYPixels = 0)
        {
            Sheet.AddImageFromUrlAt(row, column, url, widthPixels, heightPixels, offsetXPixels, offsetYPixels);
            return this;
        }
    }
}
