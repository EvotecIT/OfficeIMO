#nullable enable

using System.IO.Compression;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

/// <summary>
/// Opens CSV file readers and writers with OfficeIMO CSV encoding and compression options.
/// </summary>
public static class CsvFile
{
    /// <summary>
    /// Opens a text reader for a CSV file, applying compression from the supplied options or file extension.
    /// </summary>
    public static TextReader OpenTextReader(string path, CsvLoadOptions? options = null, int bufferSize = 256 * 1024)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        options ??= new CsvLoadOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var stream = OpenReadStream(path, options, bufferSize, useAsync: false);
        return new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: bufferSize);
    }

    internal static TextReader OpenTextReaderForAsyncRead(string path, CsvLoadOptions options, int bufferSize = 256 * 1024)
    {
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var stream = OpenReadStream(path, options, bufferSize, useAsync: true);
        return new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: bufferSize);
    }

    internal static TextReader OpenTextReader(Stream source, CsvLoadOptions options, bool leaveOpen, int bufferSize = 256 * 1024)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (!source.CanRead) throw new ArgumentException("Source stream must be readable.", nameof(source));

        CsvCompressionType compressionType = options.CompressionType == CsvCompressionType.Auto
            ? CsvCompressionType.None
            : options.CompressionType;
        EnsureCompressionSupported(compressionType);
        ValidateMaxDecompressedBytes(options);

        Stream input = WrapReadStream(source, compressionType, leaveOpen);
        if (compressionType != CsvCompressionType.None && options.MaxDecompressedBytes is { } maxBytesLimit)
        {
            bool leaveBoundedInputOpen = compressionType == CsvCompressionType.None && leaveOpen;
            input = new CsvBoundedReadStream(input, maxBytesLimit, leaveBoundedInputOpen);
        }

        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        return new StreamReader(
            input,
            encoding,
            detectEncodingFromByteOrderMarks: true,
            bufferSize,
            leaveOpen: compressionType == CsvCompressionType.None && leaveOpen);
    }

    /// <summary>
    /// Creates a text writer for a CSV file, applying compression from the supplied options or file extension.
    /// </summary>
    public static TextWriter CreateTextWriter(string path, CsvSaveOptions? options = null, bool append = false, int bufferSize = 256 * 1024)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        options ??= new CsvSaveOptions();
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var stream = CreateWriteStream(path, options, append, bufferSize);
        return new StreamWriter(stream, encoding, bufferSize: bufferSize);
    }

    internal static TextWriter CreateTextWriterForCompressionPath(
        string writePath,
        string compressionPath,
        CsvSaveOptions options,
        int bufferSize = 256 * 1024)
    {
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        var stream = CreateWriteStream(writePath, ResolveCompression(options.CompressionType, compressionPath), options.CompressionLevel, append: false, bufferSize);
        return new StreamWriter(stream, encoding, bufferSize: bufferSize);
    }

    internal static TextWriter CreateTextWriter(
        Stream destination,
        CsvSaveOptions options,
        bool leaveOpen,
        int bufferSize = 256 * 1024)
    {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));
        var compressionType = options.CompressionType == CsvCompressionType.Auto
            ? CsvCompressionType.None
            : options.CompressionType;
        EnsureCompressionSupported(compressionType);
        EnsureCompressionLevelSupported(compressionType, options.CompressionLevel);
        var encoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
        Stream output = WrapWriteStream(destination, compressionType, options.CompressionLevel, leaveOpen);
        return new StreamWriter(output, encoding, bufferSize, leaveOpen: compressionType == CsvCompressionType.None && leaveOpen);
    }

    /// <summary>
    /// Resolves an explicit or extension-inferred compression type for a CSV path.
    /// </summary>
    public static CsvCompressionType ResolveCompression(CsvCompressionType compressionType, string path)
    {
        if (compressionType != CsvCompressionType.Auto)
        {
            return compressionType;
        }

        var extension = Path.GetExtension(path);
        if (string.Equals(extension, ".gz", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(extension, ".gzip", StringComparison.OrdinalIgnoreCase))
        {
            return CsvCompressionType.GZip;
        }

        if (string.Equals(extension, ".deflate", StringComparison.OrdinalIgnoreCase))
        {
            return CsvCompressionType.Deflate;
        }

        if (string.Equals(extension, ".br", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(extension, ".brotli", StringComparison.OrdinalIgnoreCase))
        {
            return CsvCompressionType.Brotli;
        }

        if (string.Equals(extension, ".zlib", StringComparison.OrdinalIgnoreCase))
        {
            return CsvCompressionType.ZLib;
        }

        return CsvCompressionType.None;
    }

    private static Stream OpenReadStream(string path, CsvLoadOptions options, int bufferSize, bool useAsync)
    {
        var compressionType = ResolveCompression(options.CompressionType, path);
        EnsureCompressionSupported(compressionType);
        ValidateMaxDecompressedBytes(options);
        ValidateMaxInputBytes(options);

        FileOptions fileOptions = FileOptions.SequentialScan;
        if (useAsync)
        {
            fileOptions |= FileOptions.Asynchronous;
        }

        var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, fileOptions);
        Stream stream = WrapReadStream(fileStream, compressionType, leaveOpen: false);
        if (compressionType == CsvCompressionType.None)
        {
            stream = new CsvBoundedReadStream(stream, options.MaxInputBytes, leaveOpen: false);
        }
        else if (options.MaxDecompressedBytes is { } maxBytesLimit)
        {
            stream = new CsvBoundedReadStream(stream, maxBytesLimit, leaveOpen: false);
        }

        return stream;
    }

    private static void ValidateMaxInputBytes(CsvLoadOptions options)
    {
        if (options.MaxInputBytes <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options.MaxInputBytes));
        }
    }

    private static Stream CreateWriteStream(string path, CsvSaveOptions options, bool append, int bufferSize) =>
        CreateWriteStream(path, ResolveCompression(options.CompressionType, path), options.CompressionLevel, append, bufferSize);

    private static Stream CreateWriteStream(string path, CsvCompressionType compressionType, CompressionLevel compressionLevel, bool append, int bufferSize)
    {
        EnsureCompressionSupported(compressionType);
        EnsureCompressionLevelSupported(compressionType, compressionLevel);
        if (append && compressionType != CsvCompressionType.None)
        {
            throw new NotSupportedException("Appending to compressed CSV files is not supported.");
        }

        var mode = append ? FileMode.Append : FileMode.Create;
        var fileStream = new FileStream(path, mode, FileAccess.Write, FileShare.Read, bufferSize, FileOptions.SequentialScan);
        return WrapWriteStream(fileStream, compressionType, compressionLevel);
    }

    private static Stream WrapReadStream(Stream stream, CsvCompressionType compressionType, bool leaveOpen) =>
        compressionType switch
        {
            CsvCompressionType.None => stream,
            CsvCompressionType.GZip => new GZipStream(stream, CompressionMode.Decompress, leaveOpen),
            CsvCompressionType.Deflate => new DeflateStream(stream, CompressionMode.Decompress, leaveOpen),
#if NET8_0_OR_GREATER
            CsvCompressionType.Brotli => new BrotliStream(stream, CompressionMode.Decompress, leaveOpen),
            CsvCompressionType.ZLib => new ZLibStream(stream, CompressionMode.Decompress, leaveOpen),
#else
            CsvCompressionType.Brotli => throw new PlatformNotSupportedException("Brotli CSV compression requires a .NET runtime that supports BrotliStream."),
            CsvCompressionType.ZLib => throw new PlatformNotSupportedException("ZLib CSV compression requires a .NET runtime that supports ZLibStream."),
#endif
            _ => throw new ArgumentOutOfRangeException(nameof(compressionType), compressionType, "Unsupported CSV compression type.")
        };

    private static void ValidateMaxDecompressedBytes(CsvLoadOptions options)
    {
        if (options.MaxDecompressedBytes is { } maxBytes && maxBytes < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "MaxDecompressedBytes cannot be negative.");
        }
    }

    private static Stream WrapWriteStream(Stream stream, CsvCompressionType compressionType, CompressionLevel compressionLevel) =>
        WrapWriteStream(stream, compressionType, compressionLevel, leaveOpen: false);

    private static Stream WrapWriteStream(Stream stream, CsvCompressionType compressionType, CompressionLevel compressionLevel, bool leaveOpen) =>
        compressionType switch
        {
            CsvCompressionType.None => stream,
            CsvCompressionType.GZip => new GZipStream(stream, compressionLevel, leaveOpen),
            CsvCompressionType.Deflate => new DeflateStream(stream, compressionLevel, leaveOpen),
#if NET8_0_OR_GREATER
            CsvCompressionType.Brotli => new BrotliStream(stream, compressionLevel, leaveOpen),
            CsvCompressionType.ZLib => new ZLibStream(stream, compressionLevel, leaveOpen),
#else
            CsvCompressionType.Brotli => throw new PlatformNotSupportedException("Brotli CSV compression requires a .NET runtime that supports BrotliStream."),
            CsvCompressionType.ZLib => throw new PlatformNotSupportedException("ZLib CSV compression requires a .NET runtime that supports ZLibStream."),
#endif
            _ => throw new ArgumentOutOfRangeException(nameof(compressionType), compressionType, "Unsupported CSV compression type.")
        };

    private static void EnsureCompressionLevelSupported(CsvCompressionType compressionType, CompressionLevel compressionLevel)
    {
        if (compressionType == CsvCompressionType.None)
        {
            return;
        }

        if (!Enum.IsDefined(typeof(CompressionLevel), compressionLevel))
        {
            throw new ArgumentOutOfRangeException(nameof(compressionLevel), compressionLevel, "Unsupported CSV compression level.");
        }
    }

    private static void EnsureCompressionSupported(CsvCompressionType compressionType)
    {
        switch (compressionType)
        {
            case CsvCompressionType.None:
            case CsvCompressionType.GZip:
            case CsvCompressionType.Deflate:
                return;
#if NET8_0_OR_GREATER
            case CsvCompressionType.Brotli:
            case CsvCompressionType.ZLib:
                return;
#else
            case CsvCompressionType.Brotli:
                throw new PlatformNotSupportedException("Brotli CSV compression requires a .NET runtime that supports BrotliStream.");
            case CsvCompressionType.ZLib:
                throw new PlatformNotSupportedException("ZLib CSV compression requires a .NET runtime that supports ZLibStream.");
#endif
            default:
                throw new ArgumentOutOfRangeException(nameof(compressionType), compressionType, "Unsupported CSV compression type.");
        }
    }

    private sealed class CsvBoundedReadStream : Stream
    {
        private readonly Stream _inner;
        private readonly long _maxBytes;
        private readonly bool _leaveOpen;
        private long _bytesRead;

        public CsvBoundedReadStream(Stream inner, long maxBytes, bool leaveOpen)
        {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            _maxBytes = maxBytes;
            _leaveOpen = leaveOpen;
        }

        public override bool CanRead => _inner.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();

        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => _inner.Flush();

        public override int Read(byte[] buffer, int offset, int count)
        {
            var read = _inner.Read(buffer, offset, count);
            RecordBytesRead(read);
            return read;
        }

        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            int read = await _inner.ReadAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
            RecordBytesRead(read);
            return read;
        }

        private void RecordBytesRead(int count)
        {
            _bytesRead += count;
            if (_bytesRead > _maxBytes)
            {
                throw new InvalidOperationException($"CSV data exceeded the configured limit of {_maxBytes} bytes.");
            }
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing && !_leaveOpen)
            {
                _inner.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
