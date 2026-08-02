using OfficeIMO.Excel.Xlsb.Package;
using System.Buffers;
using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Read {
    internal sealed partial class XlsbTabularDataReader {
        private static XlsbPooledPartStream CreatePooledPart(
            Stream worksheetPart,
            XlsbImportOptions limits,
            CancellationToken cancellationToken) {
            if (worksheetPart == null) {
                throw new ArgumentNullException(nameof(worksheetPart));
            }
            try {
                if (!worksheetPart.CanRead || !worksheetPart.CanSeek) {
                    throw new InvalidOperationException(
                        "XLSB reads require a readable, seekable worksheet part.");
                }
                long remaining = worksheetPart.Length - worksheetPart.Position;
                if (remaining < 0 || remaining > limits.MaxPartBytes || remaining > int.MaxValue) {
                    throw new InvalidDataException(
                        $"The XLSB worksheet part contains {remaining} bytes, exceeding the configured limit of {limits.MaxPartBytes} bytes.");
                }
                int length = checked((int)remaining);
                byte[] buffer = ArrayPool<byte>.Shared.Rent(Math.Max(1, length));
                try {
                    int offset = 0;
                    while (offset < length) {
                        cancellationToken.ThrowIfCancellationRequested();
                        int read = worksheetPart.Read(buffer, offset, length - offset);
                        if (read == 0) {
                            throw new EndOfStreamException(
                                $"The XLSB worksheet part ended after {offset} of {length} declared bytes.");
                        }
                        offset += read;
                    }
                    cancellationToken.ThrowIfCancellationRequested();
                    return new XlsbPooledPartStream(buffer, length, worksheetPart);
                } catch {
                    ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
                    throw;
                }
            } catch {
                worksheetPart.Dispose();
                throw;
            }
        }
    }
}
