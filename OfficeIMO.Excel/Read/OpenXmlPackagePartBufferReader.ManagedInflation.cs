#if NET8_0_OR_GREATER
using System.Buffers;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel;

internal sealed partial class OpenXmlPackagePartBufferReader {
    /// <summary>Tries bounded memory inflation while the existing worker qualifies stable row prefixes.</summary>
    private bool TryReadManagedPipelined(ZipArchiveEntry entry, int total, CancellationToken ct,
        Func<byte[], int, WorksheetInflationProgress, bool> consume,
        out byte[]? buffer, out bool consumed) {
        buffer = null;
        consumed = false;
        // .NET 11 adds ZIP CRC validation during reads; retain its stream so that proof is not bypassed.
        if (Environment.Version.Major >= 11) return false;
        // Keep BCL inflation on hosts outside the qualified macOS ARM64 path.
        if (!OperatingSystem.IsMacOS() || RuntimeInformation.ProcessArchitecture != Architecture.Arm64) return false;
        if (!TryReadCompressedInput(entry, ct, out byte[]? compressed, out int compressedLength)) return false;
        byte[]? output = null;
        Task<bool>? consumer = null;
        var progress = new WorksheetInflationProgress();
        bool checkedPrefix = false, acceptedPrefix = false;
        try {
            output = OpenXmlPartBufferPool.Rent(total);
            byte[] bytes = output;
            bool valid = OfficeDeflateStreamValidator.TryInflateExact(compressed!, 0, compressedLength,
                output, total, written => {
                    if (!checkedPrefix) {
                        if (written < 256) return;
                        checkedPrefix = true;
                        acceptedPrefix = ExcelSheetReader.CanBufferUtf8WorksheetPrefix(bytes, 256);
                        if (acceptedPrefix)
                            consumer = Task.Run(() => consume(bytes, total, progress), CancellationToken.None);
                    }
                    if (acceptedPrefix) progress.PublishRows(bytes, written);
                }, ct);
            if (!valid) return false;
            if (!checkedPrefix) {
                acceptedPrefix = ExcelSheetReader.CanBufferUtf8WorksheetPrefix(bytes, 256);
                if (acceptedPrefix)
                    consumer = Task.Run(() => consume(bytes, total, progress), CancellationToken.None);
            }
            if (!acceptedPrefix) return false;
            progress.Complete(total);
            consumed = consumer!.GetAwaiter().GetResult();
            ct.ThrowIfCancellationRequested();
            buffer = output;
            output = null;
            return true;
        } finally {
            // A failed attempt joins its worker before BCL fallback can reuse storage.
            progress.Abort();
            if (consumer != null) {
                try { consumer.GetAwaiter().GetResult(); } catch { }
            }
            if (output != null) OpenXmlPartBufferPool.Return(output);
            ArrayPool<byte>.Shared.Return(compressed!, clearArray: true);
        }
    }

    /// <summary>Uses the shared ZIP coordinate and local-header proof without decoding the member twice.</summary>
    private bool TryReadCompressedInput(ZipArchiveEntry entry, CancellationToken ct,
        out byte[]? buffer, out int length) {
        buffer = null;
        length = 0;
        if (!_stream.CanSeek || entry.CompressedLength <= 0 || entry.CompressedLength > int.MaxValue) return false;
        byte[]? rented = null;
        long position = _stream.Position;
        try {
                if (!TryGetZipCoordinates(entry, ct, out var item)) return false;
            if (item.CompressionMethod != 8 || item.CompressedSize != entry.CompressedLength
                || item.UncompressedSize != entry.Length) return false;
            long payload = OfficeArchiveSafety.ValidateEntryPayload(_stream, item);
            int count = checked((int)item.CompressedSize);
            rented = ArrayPool<byte>.Shared.Rent(count);
            _stream.Position = payload;
            int read = 0;
            while (read < count) {
                ct.ThrowIfCancellationRequested();
                int next = _stream.Read(rented, read, Math.Min(81920, count - read));
                if (next == 0) return false;
                read += next;
            }
            buffer = rented;
            rented = null;
            length = count;
            return true;
        } catch (InvalidDataException) {
            return false;
        } catch (IOException) {
            return false;
        } catch (NotSupportedException) {
            return false;
        } catch (ArgumentException) {
            return false;
        } finally {
            _stream.Position = position;
            if (rented != null) ArrayPool<byte>.Shared.Return(rented, clearArray: true);
        }
    }
}
#endif
