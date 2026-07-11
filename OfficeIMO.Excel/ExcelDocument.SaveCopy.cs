using OfficeIMO.Shared;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Saves an independent copy and returns the workbook loaded from that copy.</summary>
        /// <param name="filePath">Destination XLSX or XLS path.</param>
        /// <param name="options">Optional save settings, including <see cref="ExcelSaveOptions.OpenAfterSave"/>.</param>
        /// <returns>A new workbook associated with <paramref name="filePath"/>. This instance keeps its current path.</returns>
        public ExcelDocument SaveCopy(string filePath, ExcelSaveOptions? options = null) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));

            string extension = Path.GetExtension(filePath);
            ExcelFileFormat format;
            if (string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)) {
                format = ExcelFileFormat.Xlsx;
            } else if (string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase)) {
                format = ExcelFileFormat.Xls;
            } else {
                throw new NotSupportedException("SaveCopy supports .xlsx and .xls destinations. Use Save for macro-enabled or template formats.");
            }

            byte[] bytes = format == ExcelFileFormat.Xls ? ToXls(options) : ToXlsx(options);
            OfficeFileCommit.WriteAllBytes(filePath, bytes);
            if (options?.OpenAfterSave == true) Open(filePath, true);
            return Load(filePath);
        }

        /// <summary>Saves an independent stream copy and returns a workbook loaded from it.</summary>
        /// <param name="outputStream">Readable, writable, seekable destination stream.</param>
        /// <param name="format">Physical XLSX or XLS format.</param>
        /// <param name="options">Optional save settings.</param>
        /// <returns>A new workbook backed by <paramref name="outputStream"/>.</returns>
        public ExcelDocument SaveCopy(Stream outputStream, ExcelFileFormat format = ExcelFileFormat.Xlsx, ExcelSaveOptions? options = null) {
            if (outputStream == null) throw new ArgumentNullException(nameof(outputStream));
            if (!outputStream.CanRead || !outputStream.CanWrite || !outputStream.CanSeek) {
                throw new ArgumentException("Stream must support reading, writing, and seeking.", nameof(outputStream));
            }

            Save(outputStream, format, options);
            outputStream.Seek(0, SeekOrigin.Begin);
            return Load(outputStream);
        }
    }
}
