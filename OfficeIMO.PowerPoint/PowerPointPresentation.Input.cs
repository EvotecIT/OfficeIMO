using OfficeIMO.Drawing.Internal;
using OfficeIMO.PowerPoint.LegacyPpt;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        private const int CompoundDetectionDirectoryLimit = 65536;
        private const int InputCopyBufferSize = 81920;

        internal static byte[] ReadPresentationInputBytes(
            Stream source,
            PowerPointLoadOptions options,
            CancellationToken cancellationToken = default) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (!source.CanRead) {
                throw new ArgumentException("Stream must be readable.",
                    nameof(source));
            }

            if (source.CanSeek) {
                long originalPosition = source.Position;
                try {
                    source.Position = 0;
                    OfficeCompoundDocumentDetector.DocumentKind kind =
                        OfficeCompoundDocumentDetector.Detect(source,
                            long.MaxValue, CompoundDetectionDirectoryLimit,
                            cancellationToken, out _);
                    long? maxInputBytes = ResolveInputLimit(kind, options);
                    return OfficeStreamReader.ReadAllBytes(source,
                        cancellationToken, maxInputBytes);
                } finally {
                    source.Position = originalPosition;
                }
            }

            byte[] prefix = ReadPrefix(source, cancellationToken);
            if (!OfficeCompoundDocumentDetector.HasCompoundSignature(prefix)) {
                return ReadRemainderToMemory(source, prefix,
                    cancellationToken);
            }
            return ReadCompoundInputThroughTemporaryStorage(source, prefix,
                options, cancellationToken);
        }

        internal static async Task<byte[]> ReadPresentationInputBytesAsync(
            Stream source,
            PowerPointLoadOptions options,
            CancellationToken cancellationToken = default) {
            if (source == null) throw new ArgumentNullException(nameof(source));
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (!source.CanRead) {
                throw new ArgumentException("Stream must be readable.",
                    nameof(source));
            }

            long originalPosition = source.CanSeek ? source.Position : 0L;
            try {
                if (source.CanSeek) source.Position = 0;
                byte[] prefix = await ReadPrefixAsync(source,
                    cancellationToken).ConfigureAwait(false);
                if (!OfficeCompoundDocumentDetector
                        .HasCompoundSignature(prefix)) {
                    return await ReadRemainderToMemoryAsync(source, prefix,
                        cancellationToken).ConfigureAwait(false);
                }
                return await ReadCompoundInputThroughTemporaryStorageAsync(
                    source, prefix, options, cancellationToken)
                    .ConfigureAwait(false);
            } finally {
                if (source.CanSeek) source.Position = originalPosition;
            }
        }

        private static byte[] ReadCompoundInputThroughTemporaryStorage(
            Stream source,
            byte[] prefix,
            PowerPointLoadOptions options,
            CancellationToken cancellationToken) {
            using FileStream temporary = CreateTemporaryInputStream(
                useAsync: false);
            temporary.Write(prefix, 0, prefix.Length);
            CopyTo(source, temporary, cancellationToken);
            temporary.Position = 0;
            OfficeCompoundDocumentDetector.DocumentKind kind =
                OfficeCompoundDocumentDetector.Detect(temporary,
                    long.MaxValue, CompoundDetectionDirectoryLimit,
                    cancellationToken, out _);
            long? maxInputBytes = ResolveInputLimit(kind, options);
            return OfficeStreamReader.ReadAllBytes(temporary,
                cancellationToken, maxInputBytes);
        }

        private static async Task<byte[]>
            ReadCompoundInputThroughTemporaryStorageAsync(
                Stream source,
                byte[] prefix,
                PowerPointLoadOptions options,
                CancellationToken cancellationToken) {
            using FileStream temporary = CreateTemporaryInputStream(
                useAsync: true);
            await temporary.WriteAsync(prefix, 0, prefix.Length,
                cancellationToken).ConfigureAwait(false);
            await CopyToAsync(source, temporary, cancellationToken)
                .ConfigureAwait(false);
            await temporary.FlushAsync(cancellationToken)
                .ConfigureAwait(false);
            temporary.Position = 0;
            OfficeCompoundDocumentDetector.DocumentKind kind =
                OfficeCompoundDocumentDetector.Detect(temporary,
                    long.MaxValue, CompoundDetectionDirectoryLimit,
                    cancellationToken, out _);
            long? maxInputBytes = ResolveInputLimit(kind, options);
            return await OfficeStreamReader.ReadAllBytesAsync(temporary,
                cancellationToken, maxInputBytes).ConfigureAwait(false);
        }

        private static long? ResolveInputLimit(
            OfficeCompoundDocumentDetector.DocumentKind kind,
            PowerPointLoadOptions options) {
            bool useLegacyLimit = kind is
                OfficeCompoundDocumentDetector.DocumentKind
                    .PowerPointPresentation
                or OfficeCompoundDocumentDetector.DocumentKind.Ambiguous
                or OfficeCompoundDocumentDetector.DocumentKind.UnknownCompound;
            if (!useLegacyLimit) return null;
            return ResolveLegacyInputLimit(options);
        }

        private static int ResolveLegacyInputLimit(
            PowerPointLoadOptions options) {
            int limit = options.LegacyPptImportOptions?.MaxInputBytes
                ?? LegacyPptImportOptions.DefaultMaxInputBytes;
            if (limit < 1) {
                throw new ArgumentOutOfRangeException(
                    nameof(LegacyPptImportOptions.MaxInputBytes));
            }
            return limit;
        }

        private static byte[] ReadPrefix(Stream source,
            CancellationToken cancellationToken) {
            var prefix = new byte[8];
            int total = 0;
            while (total < prefix.Length) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(prefix, total, prefix.Length - total);
                if (read == 0) break;
                total += read;
            }
            if (total == prefix.Length) return prefix;
            Array.Resize(ref prefix, total);
            return prefix;
        }

        private static async Task<byte[]> ReadPrefixAsync(Stream source,
            CancellationToken cancellationToken) {
            var prefix = new byte[8];
            int total = 0;
            while (total < prefix.Length) {
                int read = await source.ReadAsync(prefix, total,
                    prefix.Length - total, cancellationToken)
                    .ConfigureAwait(false);
                if (read == 0) break;
                total += read;
            }
            if (total == prefix.Length) return prefix;
            Array.Resize(ref prefix, total);
            return prefix;
        }

        private static byte[] ReadRemainderToMemory(Stream source,
            byte[] prefix, CancellationToken cancellationToken) {
            using var output = new MemoryStream();
            output.Write(prefix, 0, prefix.Length);
            CopyTo(source, output, cancellationToken);
            return output.ToArray();
        }

        private static async Task<byte[]> ReadRemainderToMemoryAsync(
            Stream source, byte[] prefix,
            CancellationToken cancellationToken) {
            using var output = new MemoryStream();
            await output.WriteAsync(prefix, 0, prefix.Length,
                cancellationToken).ConfigureAwait(false);
            await CopyToAsync(source, output, cancellationToken)
                .ConfigureAwait(false);
            return output.ToArray();
        }

        private static void CopyTo(Stream source, Stream destination,
            CancellationToken cancellationToken) {
            var buffer = new byte[InputCopyBufferSize];
            while (true) {
                cancellationToken.ThrowIfCancellationRequested();
                int read = source.Read(buffer, 0, buffer.Length);
                if (read == 0) break;
                destination.Write(buffer, 0, read);
            }
            cancellationToken.ThrowIfCancellationRequested();
        }

        private static async Task CopyToAsync(Stream source,
            Stream destination, CancellationToken cancellationToken) {
            var buffer = new byte[InputCopyBufferSize];
            while (true) {
                int read = await source.ReadAsync(buffer, 0, buffer.Length,
                    cancellationToken).ConfigureAwait(false);
                if (read == 0) break;
                await destination.WriteAsync(buffer, 0, read,
                    cancellationToken).ConfigureAwait(false);
            }
            cancellationToken.ThrowIfCancellationRequested();
        }

        private static FileStream CreateTemporaryInputStream(bool useAsync) {
            string path = Path.Combine(Path.GetTempPath(),
                "officeimo-powerpoint-" + Guid.NewGuid().ToString("N")
                + ".tmp");
            FileOptions options = FileOptions.DeleteOnClose;
            if (useAsync) options |= FileOptions.Asynchronous;
            return new FileStream(path, FileMode.CreateNew,
                FileAccess.ReadWrite, FileShare.None, InputCopyBufferSize,
                options);
        }
    }
}
