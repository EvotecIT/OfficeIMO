#nullable enable

using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>Decrypts a password-encrypted Office Open XML workbook and opens a forward-only data reader.</summary>
        /// <param name="path">Path to an encrypted XLSX, XLSM, XLTX, XLTM, XLAM, or XLSB package.</param>
        /// <param name="password">The package password.</param>
        /// <param name="options">Worksheet selection, reader limits, conversion, and cancellation settings.</param>
        /// <returns>A reader owned by the caller. The input file is closed before this method returns.</returns>
        /// <remarks>
        /// The encrypted input and the decrypted package are each bounded by ExcelReadOptions.MaxInputBytes.
        /// Opening buffers the complete input and decrypted package, verifies package integrity, and copies
        /// the options. Only Office Agile encryption is supported. Legacy XLS password encryption uses LoadEncrypted.
        /// </remarks>
        public static ExcelWorkbookDataReader OpenEncryptedDataReader(
            string path, string password, ExcelReadOptions? options = null) {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
            if (password == null) throw new ArgumentNullException(nameof(password));
            ExcelReadOptions effectiveOptions = (options ?? new ExcelReadOptions()).Clone();
            effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
            using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            byte[] encryptedBytes = OfficeStreamReader.ReadRemainingBytes(
                source, effectiveOptions.CancellationToken, effectiveOptions.MaxInputBytes);
            return OpenEncryptedDataReaderCore(encryptedBytes, password, effectiveOptions);
        }

        /// <summary>Decrypts the remaining encrypted Office Open XML workbook bytes and opens a forward-only data reader.</summary>
        /// <param name="stream">Readable stream positioned at the encrypted workbook bytes.</param>
        /// <param name="password">The package password.</param>
        /// <param name="options">Worksheet selection, reader limits, conversion, and cancellation settings.</param>
        /// <returns>A reader owned by the caller.</returns>
        /// <remarks>
        /// The source remains open on success, failure, and reader disposal. Its seekable position is restored.
        /// Non-seekable sources advance through their remaining content. The complete encrypted input and
        /// decrypted package are buffered, each bounded by ExcelReadOptions.MaxInputBytes. Agile package integrity
        /// is verified and options are copied. This API supports Office Open XML packages protected by Agile encryption.
        /// </remarks>
        public static ExcelWorkbookDataReader OpenEncryptedDataReader(
            Stream stream, string password, ExcelReadOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            ExcelReadOptions effectiveOptions = (options ?? new ExcelReadOptions()).Clone();
            long originalPosition = stream.CanSeek ? stream.Position : 0L;
            byte[] encryptedBytes;
            try {
                encryptedBytes = OfficeStreamReader.ReadRemainingBytes(
                    stream, effectiveOptions.CancellationToken, effectiveOptions.MaxInputBytes);
            } finally {
                if (stream.CanSeek) stream.Position = originalPosition;
            }
            return OpenEncryptedDataReaderCore(encryptedBytes, password, effectiveOptions);
        }

        /// <summary>Decrypts an in-memory Office Open XML workbook and opens a forward-only data reader.</summary>
        /// <param name="encryptedBytes">The complete encrypted Office package.</param>
        /// <param name="password">The package password.</param>
        /// <param name="options">Worksheet selection, reader limits, conversion, and cancellation settings.</param>
        /// <returns>A reader owned by the caller.</returns>
        /// <remarks>
        /// Encrypted input and decrypted package sizes are each bounded by ExcelReadOptions.MaxInputBytes.
        /// The complete decrypted package is buffered, integrity is verified, and options are copied.
        /// Only Office Agile encryption is supported.
        /// Legacy XLS password encryption is handled by LoadEncrypted.
        /// </remarks>
        public static ExcelWorkbookDataReader OpenEncryptedDataReader(
            byte[] encryptedBytes, string password, ExcelReadOptions? options = null) {
            if (encryptedBytes == null) throw new ArgumentNullException(nameof(encryptedBytes));
            if (password == null) throw new ArgumentNullException(nameof(password));
            return OpenEncryptedDataReaderCore(
                encryptedBytes, password, (options ?? new ExcelReadOptions()).Clone());
        }

        private static ExcelWorkbookDataReader OpenEncryptedDataReaderCore(
            byte[] encryptedBytes, string password, ExcelReadOptions options) {
            options.CancellationToken.ThrowIfCancellationRequested();
            if (encryptedBytes.LongLength > options.MaxInputBytes) {
                throw new InvalidDataException(
                    $"Encrypted workbook input contains {encryptedBytes.LongLength} bytes, exceeding the configured limit of {options.MaxInputBytes} bytes.");
            }
            byte[] packageBytes = OfficeEncryption.DecryptPackage(
                encryptedBytes, password, options.CancellationToken, options.MaxInputBytes);
            options.CancellationToken.ThrowIfCancellationRequested();
            if (packageBytes.Length < 2 || packageBytes[0] != 0x50 || packageBytes[1] != 0x4B) {
                throw new InvalidDataException("The decrypted workbook must be an Office Open XML package. Legacy XLS password encryption uses LoadEncrypted.");
            }
            return OpenDataReader(packageBytes, options);
        }
    }
}
