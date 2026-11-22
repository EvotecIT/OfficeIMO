using System.IO.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Visio {
    // Load orchestrator for VisioDocument.
    public partial class VisioDocument {
        /// <summary>
        /// Loads an existing .vsdx file into a VisioDocument.
        /// </summary>
        public static VisioDocument Load(string filePath) => Load(filePath, options: null);

        /// <summary>Loads an existing .vsdx file with explicit input limits.</summary>
        public static VisioDocument Load(string filePath, VisioLoadOptions? options) {
            ArgumentNullException.ThrowIfNull(filePath);
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath)) throw new FileNotFoundException($"File '{fullPath}' doesn't exist.", fullPath);

            VisioLoadOptions resolved = options ?? new VisioLoadOptions();
            using var source = OpenLoadFile(fullPath, resolved, useAsync: false, out bool canBorrow);
            if (canBorrow && CanLoadInPlace(source, resolved)) {
                return LoadDocumentInPlace(source, fullPath, sourceStream: null, resolved);
            }
            byte[] bytes = OfficeStreamReader.ReadAllBytes(source, ResolveInputLimit(resolved));
            return LoadDocument(bytes, fullPath, sourceStream: null, options: resolved);
        }

        /// <summary>
        /// Loads an existing .vsdx document from a stream.
        /// </summary>
        public static VisioDocument Load(Stream stream) => Load(stream, options: null);

        /// <summary>Loads an existing .vsdx document from a stream with explicit input limits.</summary>
        /// <param name="stream">Readable source. The caller keeps ownership and disposes it.</param>
        /// <param name="options">Input limits and package-security options.</param>
        /// <remarks>
        /// A seekable Open XML stream is inspected and read in place. Keep it unchanged and give the loader
        /// exclusive use until it returns. The caller's position is restored and the diagram is fully loaded by
        /// then. Non-seekable streams are buffered. A writable, seekable stream becomes the destination of
        /// <see cref="Save()"/>.
        /// </remarks>
        public static VisioDocument Load(Stream stream, VisioLoadOptions? options) {
            ArgumentNullException.ThrowIfNull(stream);
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            VisioLoadOptions resolved = options ?? new VisioLoadOptions();
            Stream? associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(
                stream,
                DocumentAccessMode.ReadWrite);
            if (CanLoadInPlace(stream, resolved)) {
                return LoadDocumentInPlace(stream, filePath: null, associatedStream, resolved);
            }
            byte[] bytes = OfficeStreamReader.ReadAllBytes(stream, ResolveInputLimit(resolved));
            return LoadDocument(
                bytes,
                filePath: null,
                sourceStream: associatedStream,
                options: resolved);
        }

        /// <summary>Asynchronously loads an existing .vsdx file with optional input limits.</summary>
        public static async Task<VisioDocument> LoadAsync(
            string filePath,
            VisioLoadOptions? options = null,
            CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(filePath);
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath)) throw new FileNotFoundException($"File '{fullPath}' doesn't exist.", fullPath);
            VisioLoadOptions resolved = options ?? new VisioLoadOptions();
            using var source = OpenLoadFile(fullPath, resolved, useAsync: true, out bool canBorrow);
            if (canBorrow && await CanLoadInPlaceAsync(source, resolved, cancellationToken).ConfigureAwait(false)) {
                return await OfficePackageLoadSource.LoadInPlaceAsync(source,
                    () => LoadDocumentInPlace(source, fullPath, sourceStream: null, resolved, cancellationToken),
                    cancellationToken).ConfigureAwait(false);
            }
            byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(source, cancellationToken, ResolveInputLimit(resolved)).ConfigureAwait(false);
            return LoadDocument(bytes, fullPath, sourceStream: null, options: resolved, cancellationToken: cancellationToken);
        }

        /// <summary>Asynchronously loads an existing .vsdx document from a caller-owned stream with optional input limits.</summary>
        /// <remarks>Stream handling follows <see cref="Load(Stream, VisioLoadOptions?)"/>. The async
        /// in-place path applies to standard <see cref="MemoryStream"/> and <see cref="FileStream"/> sources.
        /// Custom streams are buffered.</remarks>
        public static async Task<VisioDocument> LoadAsync(
            Stream stream,
            VisioLoadOptions? options = null,
            CancellationToken cancellationToken = default) {
            ArgumentNullException.ThrowIfNull(stream);
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            VisioLoadOptions resolved = options ?? new VisioLoadOptions();
            if (await CanLoadInPlaceAsync(stream, resolved, cancellationToken).ConfigureAwait(false)) {
                Stream? associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(stream, DocumentAccessMode.ReadWrite);
                return await OfficePackageLoadSource.LoadInPlaceAsync(stream,
                    () => LoadDocumentInPlace(stream, filePath: null, associatedStream, resolved, cancellationToken),
                    cancellationToken).ConfigureAwait(false);
            }
            byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, ResolveInputLimit(resolved)).ConfigureAwait(false);
            return LoadDocument(
                bytes,
                filePath: null,
                sourceStream: OfficeDocumentLifecycle.ResolveAssociatedDestination(
                    stream,
                    DocumentAccessMode.ReadWrite),
                options: resolved,
                cancellationToken: cancellationToken);
        }

        private static VisioDocument LoadDocument(
            byte[] bytes,
            string? filePath,
            Stream? sourceStream,
            VisioLoadOptions options,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            ValidatePackageSecurity(bytes, options);
            cancellationToken.ThrowIfCancellationRequested();
            using var buffer = new MemoryStream(bytes, writable: false);
            using Package package = Package.Open(buffer, FileMode.Open, FileAccess.Read);
            VisioDocument document = LoadCore(package, filePath, cancellationToken, (options.Limits ?? throw new ArgumentNullException(nameof(options.Limits))).MaxPackageXmlCharacters);
            document._loadedSignatureInfo = InspectPackageSignatures(bytes);
            document._sourceStream = sourceStream;
            cancellationToken.ThrowIfCancellationRequested();
            return document;
        }

        // A secure load borrows a file only while it can exclude writers. Keep accepting a shared
        // file when a competing handle prevents that open, by checking and reading one snapshot.
        private static FileStream OpenLoadFile(string path, VisioLoadOptions options, bool useAsync, out bool canBorrow) {
            const FileShare sharedWriters = FileShare.ReadWrite | FileShare.Delete;
            const int sharingViolation = unchecked((int)0x80070020);
            int bufferSize = useAsync ? 81920 : 4096;
            canBorrow = true;
            if (options.PackageSecurity != null) {
                try {
                    return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete, bufferSize, useAsync);
                } catch (IOException exception) when (exception.HResult == sharingViolation) {
                    canBorrow = false;
                }
            }
            return new FileStream(path, FileMode.Open, FileAccess.Read, sharedWriters, bufferSize, useAsync);
        }

        // The diagram is materialized while the package is open, so a seekable source can be read in place.
        private static bool CanLoadInPlace(Stream stream, VisioLoadOptions options) =>
            OfficePackageLoadSource.CanReadInPlace(stream, DocumentAccessMode.ReadOnly, ResolveInputLimit(options));

        private static Task<bool> CanLoadInPlaceAsync(Stream stream, VisioLoadOptions options,
            CancellationToken cancellationToken) =>
            OfficePackageLoadSource.CanReadInPlaceAsync(stream, DocumentAccessMode.ReadOnly, ResolveInputLimit(options), cancellationToken);

        private static VisioDocument LoadDocumentInPlace(
            Stream stream,
            string? filePath,
            Stream? sourceStream,
            VisioLoadOptions options,
            CancellationToken cancellationToken = default) {
            cancellationToken.ThrowIfCancellationRequested();
            long originalPosition = stream.Position;
            try {
                using OfficePackageSourceView view = OfficePackageLoadSource.OpenInPlace(stream);
                OfficePackageLoadSource.InspectInPlace(view, options.PackageSecurity, capture: false, cancellationToken);
                VisioDocument document;
                using (Package package = Package.Open(view, FileMode.Open, FileAccess.Read)) {
                    document = LoadCore(package, filePath, cancellationToken, (options.Limits ?? throw new ArgumentNullException(nameof(options.Limits))).MaxPackageXmlCharacters);
                }
                document._loadedSignatureInfo = InspectPackageSignatures(view);
                document._sourceStream = sourceStream;
                cancellationToken.ThrowIfCancellationRequested();
                return document;
            } finally {
                stream.Position = originalPosition;
            }
        }

        private static long? ResolveInputLimit(VisioLoadOptions options) {
            long? configured = (options.Limits ?? throw new ArgumentNullException(nameof(options.Limits))).MaxInputBytes;
            if (configured.HasValue && configured.Value < 1) {
                throw new ArgumentOutOfRangeException(nameof(VisioLoadLimits.MaxInputBytes));
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
