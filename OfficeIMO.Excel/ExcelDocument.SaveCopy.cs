using OfficeIMO.Core.Internal;
using DocumentFormat.OpenXml;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Saves an independent copy without changing this workbook's associated destination.</summary>
        /// <param name="filePath">Destination XLSX, XLSM, XLTX, XLTM, XLAM, XLS, or XLSB path.</param>
        /// <param name="options">Optional save policy settings.</param>
        public void SaveCopy(string filePath, ExcelSaveOptions? options = null) {
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

            EnsureEarlyBiffOutputPath(filePath);
            ExcelFileFormat format = GetSaveCopyFormat(filePath);

            EnsureDestinationFileWritable(filePath);
            EnsureDirectoryWritable(filePath);
            byte[] bytes = PrepareSaveCopyBytes(filePath, format, options);
            OfficeFileCommit.WriteAllBytes(filePath, bytes);
        }

        /// <summary>Asynchronously saves an independent copy without changing this workbook's associated destination.</summary>
        public async Task SaveCopyAsync(
            string filePath,
            ExcelSaveOptions? options = null,
            CancellationToken cancellationToken = default) {
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
            EnsureEarlyBiffOutputPath(filePath);
            ExcelFileFormat format = GetSaveCopyFormat(filePath);
            EnsureDestinationFileWritable(filePath);
            EnsureDirectoryWritable(filePath);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] bytes = PrepareSaveCopyBytes(filePath, format, options);
            await OfficeFileCommit.WriteAllBytesAsync(filePath, bytes, cancellationToken: cancellationToken)
                .ConfigureAwait(false);
        }

        private static ExcelFileFormat GetSaveCopyFormat(string filePath) {
            string extension = Path.GetExtension(filePath);
            if (string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)) return ExcelFileFormat.Xlsx;
            if (string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase)) return ExcelFileFormat.Xlsx;
            if (string.Equals(extension, ".xltx", StringComparison.OrdinalIgnoreCase)) return ExcelFileFormat.Xlsx;
            if (string.Equals(extension, ".xltm", StringComparison.OrdinalIgnoreCase)) return ExcelFileFormat.Xlsx;
            if (string.Equals(extension, ".xlam", StringComparison.OrdinalIgnoreCase)) return ExcelFileFormat.Xlsx;
            if (string.Equals(extension, ".xls", StringComparison.OrdinalIgnoreCase)) return ExcelFileFormat.Xls;
            if (string.Equals(extension, ".xlsb", StringComparison.OrdinalIgnoreCase)) return ExcelFileFormat.Xlsb;
            throw new NotSupportedException("SaveCopy supports .xlsx, .xlsm, .xltx, .xltm, .xlam, .xls, and .xlsb destinations.");
        }

        private byte[] PrepareSaveCopyBytes(string filePath, ExcelFileFormat format, ExcelSaveOptions? options) {
            if (format != ExcelFileFormat.Xlsx) {
                return ToBytes(format, options);
            }

            EnsureWritableForSave();
            SpreadsheetDocumentType originalType = _spreadSheetDocument.DocumentType;
            bool containsMacroContent = HasMacros
                || originalType == SpreadsheetDocumentType.MacroEnabledWorkbook
                || originalType == SpreadsheetDocumentType.MacroEnabledTemplate;
            EnsureMacroCompatibleCopy(containsMacroContent, filePath);
            try {
                AlignSpreadsheetDocumentTypeWithFilePath(filePath);
                return ToBytes(format, options);
            } finally {
                if (_spreadSheetDocument.DocumentType != originalType) {
                    ChangeSpreadsheetDocumentType(originalType);
                    MarkPackageDirty();
                }
            }
        }
    }
}
