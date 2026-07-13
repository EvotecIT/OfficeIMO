using System;
using System.IO;
using System.IO.Packaging;
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
        public static VisioDocument Load(string filePath) => LoadCore(filePath);

        /// <summary>
        /// Loads an existing .vsdx document from a stream.
        /// </summary>
        public static VisioDocument Load(Stream stream) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            using var buffer = new MemoryStream(OfficeStreamReader.ReadAllBytes(stream), writable: false);

            using Package package = Package.Open(buffer, FileMode.Open, FileAccess.Read);
            VisioDocument document = LoadCore(package, filePath: null);
            document._sourceStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(
                stream,
                OfficeIMO.Drawing.DocumentAccessMode.ReadWrite);
            return document;
        }

        /// <summary>Asynchronously loads an existing .vsdx file.</summary>
        public static async Task<VisioDocument> LoadAsync(string filePath, CancellationToken cancellationToken = default) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            string fullPath = Path.GetFullPath(filePath);
            if (!File.Exists(fullPath)) throw new FileNotFoundException($"File '{fullPath}' doesn't exist.", fullPath);
            using var source = new FileStream(fullPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 81920, useAsync: true);
            byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(source, cancellationToken).ConfigureAwait(false);
            using var buffer = new MemoryStream(bytes, writable: false);
            using Package package = Package.Open(buffer, FileMode.Open, FileAccess.Read);
            return LoadCore(package, fullPath);
        }

        /// <summary>Asynchronously loads an existing .vsdx document from a caller-owned stream.</summary>
        public static async Task<VisioDocument> LoadAsync(Stream stream, CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            byte[] bytes = await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken).ConfigureAwait(false);
            using var buffer = new MemoryStream(bytes, writable: false);
            using Package package = Package.Open(buffer, FileMode.Open, FileAccess.Read);
            VisioDocument document = LoadCore(package, filePath: null);
            document._sourceStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(
                stream,
                OfficeIMO.Drawing.DocumentAccessMode.ReadWrite);
            return document;
        }
    }
}
