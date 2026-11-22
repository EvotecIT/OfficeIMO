namespace OfficeIMO.Core.Internal {
    /// <summary>
    /// Decides which stream an Office package is opened over when a document is loaded from caller input, so that
    /// loading holds no more copies of the file than the access mode needs.
    /// </summary>
    /// <remarks>
    /// <list type="bullet">
    /// <item><description>A read-only load of a seekable ZIP package with no package-security policy reads the
    /// package in place from the caller's stream. The caller keeps the stream open and unchanged while the document
    /// is in use, and the document never writes to it or disposes it.</description></item>
    /// <item><description>Any other read-only load opens the package over the bytes already read, through a stream
    /// that cannot write.</description></item>
    /// <item><description>An editable load opens the package over the bytes already read through a copy-on-write
    /// stream, so the bytes are copied only when the package is first written, typically when it is saved.</description></item>
    /// </list>
    /// </remarks>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    internal static class OfficePackageLoadSource {
        private const int ZipSignatureLength = 4;

        /// <summary>
        /// Creates the stream a package is opened over from bytes the loader owns, without copying them. The array
        /// is never written through the returned stream.
        /// </summary>
        public static Stream CreatePackageStream(byte[] bytes, DocumentAccessMode accessMode) {
            ArgumentNullException.ThrowIfNull(bytes);
            return accessMode == DocumentAccessMode.ReadOnly
                ? new MemoryStream(bytes, 0, bytes.Length, writable: false, publiclyVisible: true)
                : new OfficeCopyOnWriteStream(bytes);
        }

        /// <summary>
        /// Returns true when a read-only load can open the package straight over <paramref name="source"/>: the
        /// stream is readable and seekable, starts with a ZIP local file header, fits the input limit, and no
        /// package-security inspection needs the encoded bytes. The stream position is left unchanged.
        /// </summary>
        public static bool CanReadInPlace(
            Stream source,
            DocumentAccessMode accessMode,
            OfficePackageSecurityOptions? packageSecurity,
            bool captureSourcePackageSecurity,
            long? maxInputBytes) {
            if (!IsInPlaceCandidate(source, accessMode, packageSecurity, captureSourcePackageSecurity, maxInputBytes)) {
                return false;
            }

            long originalPosition = source.Position;
            try {
                source.Position = 0;
                var signature = new byte[ZipSignatureLength];
                int read = 0;
                while (read < signature.Length) {
                    int count = source.Read(signature, read, signature.Length - read);
                    if (count == 0) return false;
                    read += count;
                }
                return HasZipSignature(signature);
            } finally {
                source.Position = originalPosition;
            }
        }

        /// <summary>Probes sources with known synchronous read behavior for the asynchronous in-place path.</summary>
        public static async Task<bool> CanReadInPlaceAsync(
            Stream source,
            DocumentAccessMode accessMode,
            OfficePackageSecurityOptions? packageSecurity,
            bool captureSourcePackageSecurity,
            long? maxInputBytes,
            CancellationToken cancellationToken) {
            // ZIP and OPC readers require synchronous reads. Custom streams may implement only ReadAsync,
            // and subclasses of the built-in streams can override their synchronous behavior.
            Type sourceType = source.GetType();
            if (sourceType != typeof(FileStream) && sourceType != typeof(MemoryStream)) return false;
            if (!IsInPlaceCandidate(source, accessMode, packageSecurity, captureSourcePackageSecurity, maxInputBytes)) {
                return false;
            }

            long originalPosition = source.Position;
            try {
                source.Position = 0;
                var signature = new byte[ZipSignatureLength];
                int read = 0;
                while (read < signature.Length) {
                    int count = await source.ReadAsync(signature, read, signature.Length - read, cancellationToken)
                        .ConfigureAwait(false);
                    if (count == 0) return false;
                    read += count;
                }
                return HasZipSignature(signature);
            } finally {
                source.Position = originalPosition;
            }
        }

        /// <summary>
        /// Runs an asynchronous load's in-place work. The package and signature readers only read synchronously, so a
        /// <see cref="FileStream"/> is read on a thread-pool thread. An in-memory source is read directly.
        /// </summary>
        public static Task<T> LoadInPlaceAsync<T>(Stream source, Func<T> load, CancellationToken cancellationToken) {
            ArgumentNullException.ThrowIfNull(source);
            ArgumentNullException.ThrowIfNull(load);
            if (source.GetType() == typeof(MemoryStream)) {
                cancellationToken.ThrowIfCancellationRequested();
                return Task.FromResult(load());
            }
            return Task.Run(load, cancellationToken);
        }

        private static bool IsInPlaceCandidate(
            Stream source,
            DocumentAccessMode accessMode,
            OfficePackageSecurityOptions? packageSecurity,
            bool captureSourcePackageSecurity,
            long? maxInputBytes) {
            ArgumentNullException.ThrowIfNull(source);
            if (accessMode != DocumentAccessMode.ReadOnly
                || packageSecurity != null
                || captureSourcePackageSecurity
                || !source.CanRead
                || !source.CanSeek) {
                return false;
            }

            long length = source.Length;
            return length >= ZipSignatureLength && (!maxInputBytes.HasValue || length <= maxInputBytes.Value);
        }

        /// <summary>
        /// Wraps the caller's stream for an in-place read-only package. The view cannot write and does not dispose
        /// the caller's stream.
        /// </summary>
        public static OfficePackageSourceView OpenInPlace(Stream source) => new OfficePackageSourceView(source);

        /// <summary>Returns true when <paramref name="bytes"/> starts with a ZIP local file header.</summary>
        public static bool HasZipSignature(byte[] bytes) =>
            bytes != null
            && bytes.Length >= ZipSignatureLength
            && bytes[0] == 0x50 && bytes[1] == 0x4B && bytes[2] == 0x03 && bytes[3] == 0x04;

        /// <summary>
        /// Returns a snapshot of the encoded package held by an in-memory package stream, or false when the
        /// package is read from elsewhere.
        /// </summary>
        public static bool TryGetPackageBytes(Stream? packageStream, out byte[] bytes) {
            switch (packageStream) {
                case MemoryStream memory:
                    bytes = memory.ToArray();
                    return true;
                case OfficeCopyOnWriteStream copyOnWrite:
                    bytes = copyOnWrite.ToArray();
                    return true;
                default:
                    bytes = Array.Empty<byte>();
                    return false;
            }
        }

        /// <summary>
        /// Returns the encoded package a package stream holds or reads from: a snapshot of an in-memory stream, or
        /// the whole caller's stream for an in-place package, read without moving its position.
        /// </summary>
        public static byte[] ReadPackageBytes(Stream packageStream) {
            ArgumentNullException.ThrowIfNull(packageStream);
            return TryGetPackageBytes(packageStream, out byte[] bytes)
                ? bytes
                : OfficeStreamReader.ReadAllBytes(packageStream);
        }

        /// <summary>Stops a closing editable package from copying bytes only to rewrite a stream nothing reads again.</summary>
        public static void AbandonPackageStream(Stream? packageStream) {
            if (packageStream is OfficeCopyOnWriteStream copyOnWrite) copyOnWrite.Abandon();
        }
    }

    /// <summary>
    /// A read-only, non-owning view of a caller's seekable stream that an in-place package reads from.
    /// </summary>
    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
    internal sealed class OfficePackageSourceView : Stream {
        private readonly Stream _source;
        private bool _disposed;

        internal OfficePackageSourceView(Stream source) {
            _source = source ?? throw new ArgumentNullException(nameof(source));
            if (!source.CanRead || !source.CanSeek) {
                throw new ArgumentException("Stream must be readable and seekable.", nameof(source));
            }
        }

        public override bool CanRead => !_disposed && _source.CanRead;
        public override bool CanSeek => !_disposed && _source.CanSeek;
        public override bool CanWrite => false;
        public override long Length { get { ThrowIfDisposed(); return _source.Length; } }

        public override long Position {
            get { ThrowIfDisposed(); return _source.Position; }
            set { ThrowIfDisposed(); _source.Position = value; }
        }

        public override void Flush() { }

        public override int Read(byte[] buffer, int offset, int count) {
            ThrowIfDisposed();
            return _source.Read(buffer, offset, count);
        }

        public override int ReadByte() {
            ThrowIfDisposed();
            return _source.ReadByte();
        }

#if NET5_0_OR_GREATER
        public override int Read(Span<byte> buffer) {
            ThrowIfDisposed();
            return _source.Read(buffer);
        }
#endif

        public override long Seek(long offset, SeekOrigin origin) {
            ThrowIfDisposed();
            return _source.Seek(offset, origin);
        }

        public override void SetLength(long value) =>
            throw new NotSupportedException("A read-only document cannot change the stream it was loaded from.");

        public override void Write(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException("A read-only document cannot write to the stream it was loaded from.");

        protected override void Dispose(bool disposing) {
            // The caller owns the stream; the document only stops reading it.
            _disposed = true;
            base.Dispose(disposing);
        }

        private void ThrowIfDisposed() {
            if (_disposed) throw new ObjectDisposedException(nameof(OfficePackageSourceView));
        }
    }
}
