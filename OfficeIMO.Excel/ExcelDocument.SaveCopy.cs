using OfficeIMO.Drawing.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Saves an independent copy without changing this workbook's associated destination.</summary>
        /// <param name="filePath">Destination XLSX or XLS path.</param>
        /// <param name="options">Optional save policy settings.</param>
        public void SaveCopy(string filePath, ExcelSaveOptions? options = null) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));

            ExcelFileFormat format = GetSaveCopyFormat(filePath);

            EnsureDestinationFileWritable(filePath);
            EnsureDirectoryWritable(filePath);
            byte[] bytes = ToBytes(format, options);
            OfficeFileCommit.WriteAllBytes(filePath, bytes);
        }

        /// <summary>Asynchronously saves an independent copy without changing this workbook's associated destination.</summary>
        public async Task SaveCopyAsync(
            string filePath,
            ExcelSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            ExcelFileFormat format = GetSaveCopyFormat(filePath);
            EnsureDestinationFileWritable(filePath);
            EnsureDirectoryWritable(filePath);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes = ToBytes(format, options);
            await OfficeFileCommit.WriteAllBytesAsync(filePath, bytes, cancellationToken: cancellationToken)
                .ConfigureAwait(false);
        }

        private static ExcelFileFormat GetSaveCopyFormat(string filePath) {
            string extension = Path.GetExtension(filePath);
            if (string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)) return ExcelFileFormat.Xlsx;
            if (string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase)) return ExcelFileFormat.Xls;
            throw new NotSupportedException("SaveCopy supports .xlsx and .xls destinations. Use Save for macro-enabled or template formats.");
        }
    }
}
