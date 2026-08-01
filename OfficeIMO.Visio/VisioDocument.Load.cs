using System;
using System.IO;
using System.IO.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.Internal;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Load orchestrator for VisioDocument.
    /// </summary>
    public partial class VisioDocument {
        /// <summary>
        /// Loads an existing .vsdx file into a VisioDocument.
        /// </summary>
        public static VisioDocument Load(string filePath) => Load(filePath, options: null);

        /// <summary>Loads an existing .vsdx file with explicit input limits.</summary>
        public static VisioDocument Load(string filePath, VisioLoadOptions? options) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath)) throw new FileNotFoundException($"File '{fullPath}' doesn't exist.", fullPath);

            VisioLoadOptions resolved = options ?? new VisioLoadOptions();
            using var source = new FileStream(
                fullPath,
                FileMode.Open,
                FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete);
            byte[] bytes = OfficeStreamReader.ReadAllBytes(source, ResolveInputLimit(resolved));
            return LoadDocument(bytes, fullPath, sourceStream: null, options: resolved);
        }

        /// <summary>
        /// Loads an existing .vsdx document from a stream.
        /// </summary>
        public static VisioDocument Load(Stream stream) => Load(stream, options: null);

        /// <summary>Loads an existing .vsdx document from a stream with explicit input limits.</summary>
        public static VisioDocument Load(Stream stream, VisioLoadOptions? options) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            VisioLoadOptions resolved = options ?? new VisioLoadOptions();
            byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, ResolveInputLimit(resolved));
            return LoadDocument(
                bytes,
                filePath: null,
                sourceStream: OfficeDocumentLifecycle.ResolveAssociatedDestination(
                    stream,
                    DocumentAccessMode.ReadWrite),
                options: resolved);
        }

        /// <summary>Asynchronously loads an existing .vsdx file with optional input limits.</summary>
        public static Task<VisioDocument> LoadAsync(
            string filePath,
            CancellationToken cancellationToken = default) =>
            LoadAsync(filePath, options: null, cancellationToken);

        /// <summary>Asynchronously loads an existing .vsdx file with explicit input limits.</summary>
        public static Task<VisioDocument> LoadAsync(
            string filePath,
            CancellationToken cancellationToken,
            VisioLoadOptions? options) =>
            LoadAsync(filePath, options, cancellationToken);

        /// <summary>Asynchronously loads an existing .vsdx file with optional input limits.</summary>
        public static async Task<VisioDocument> LoadAsync(
            string filePath,
            VisioLoadOptions? options,
            CancellationToken cancellationToken = default) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath)) throw new FileNotFoundException($"File '{fullPath}' doesn't exist.", fullPath);
            using var source = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 81920, useAsync: true);
            VisioLoadOptions resolved = options ?? new VisioLoadOptions();
            byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(source, cancellationToken, ResolveInputLimit(resolved)).ConfigureAwait(false);
            return LoadDocument(bytes, fullPath, sourceStream: null, options: resolved);
        }

        /// <summary>Asynchronously loads an existing .vsdx document from a caller-owned stream with optional input limits.</summary>
        public static Task<VisioDocument> LoadAsync(
            Stream stream,
            CancellationToken cancellationToken = default) =>
            LoadAsync(stream, options: null, cancellationToken);

        /// <summary>Asynchronously loads an existing .vsdx document from a caller-owned stream with explicit input limits.</summary>
        public static Task<VisioDocument> LoadAsync(
            Stream stream,
            CancellationToken cancellationToken,
            VisioLoadOptions? options) =>
            LoadAsync(stream, options, cancellationToken);

        /// <summary>Asynchronously loads an existing .vsdx document from a caller-owned stream with optional input limits.</summary>
        public static async Task<VisioDocument> LoadAsync(
            Stream stream,
            VisioLoadOptions? options,
            CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            VisioLoadOptions resolved = options ?? new VisioLoadOptions();
            byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, ResolveInputLimit(resolved)).ConfigureAwait(false);
            return LoadDocument(
                bytes,
                filePath: null,
                sourceStream: OfficeDocumentLifecycle.ResolveAssociatedDestination(
                    stream,
                    DocumentAccessMode.ReadWrite),
                options: resolved);
        }

        private static VisioDocument LoadDocument(
            byte[] bytes,
            string? filePath,
            Stream? sourceStream,
            VisioLoadOptions options) {
            ValidatePackageSecurity(bytes, options);
            using var buffer = new MemoryStream(bytes, writable: false);
            using Package package = Package.Open(buffer, FileMode.Open, FileAccess.Read);
            VisioDocument document = LoadCore(package, filePath);
            document._sourceStream = sourceStream;
            return document;
        }

        private static long? ResolveInputLimit(VisioLoadOptions options) {
            long? configured = options.MaxInputBytes;
            if (configured.HasValue && configured.Value < 1) {
                throw new ArgumentOutOfRangeException(nameof(options.MaxInputBytes));
            }
            if (options.PackageSecurity == null) return configured;
            long packageLimit = options.PackageSecurity.MaxPackageBytes;
            if (packageLimit < 1) {
                throw new ArgumentOutOfRangeException(nameof(OfficePackageSecurityOptions.MaxPackageBytes));
            }
            return configured.HasValue ? Math.Min(configured.Value, packageLimit) : packageLimit;
        }

        private static void ValidatePackageSecurity(byte[] bytes, VisioLoadOptions options) {
            if (options.PackageSecurity != null) {
                OfficePackageSecurityInspector.Validate(bytes, options.PackageSecurity);
            }
        }
    }
}
