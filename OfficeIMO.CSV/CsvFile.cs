#nullable enable

using System.IO.Compression;
using System.Text;

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
        var stream = OpenReadStream(path, options, bufferSize);
        return new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: bufferSize);
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

    private static Stream OpenReadStream(string path, CsvLoadOptions options, int bufferSize)
    {
        var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize, FileOptions.SequentialScan);
        Stream stream = WrapReadStream(fileStream, ResolveCompression(options.CompressionType, path));
        if (options.MaxDecompressedBytes is { } maxBytes)
        {
            if (maxBytes < 0)
            {
                fileStream.Dispose();
                throw new ArgumentOutOfRangeException(nameof(options), "MaxDecompressedBytes cannot be negative.");
            }

            stream = new CsvBoundedReadStream(stream, maxBytes);
        }

        return stream;
    }

    private static Stream CreateWriteStream(string path, CsvSaveOptions options, bool append, int bufferSize) =>
        CreateWriteStream(path, ResolveCompression(options.CompressionType, path), options.CompressionLevel, append, bufferSize);

    private static Stream CreateWriteStream(string path, CsvCompressionType compressionType, CompressionLevel compressionLevel, bool append, int bufferSize)
    {
        if (append && compressionType != CsvCompressionType.None)
        {
            throw new NotSupportedException("Appending to compressed CSV files is not supported.");
        }

        var mode = append ? FileMode.Append : FileMode.Create;
        var fileStream = new FileStream(path, mode, FileAccess.Write, FileShare.Read, bufferSize, FileOptions.SequentialScan);
        return WrapWriteStream(fileStream, compressionType, compressionLevel);
    }

    private static Stream WrapReadStream(Stream stream, CsvCompressionType compressionType) =>
        compressionType switch
        {
            CsvCompressionType.None => stream,
            CsvCompressionType.GZip => new GZipStream(stream, CompressionMode.Decompress, leaveOpen: false),
            CsvCompressionType.Deflate => new DeflateStream(stream, CompressionMode.Decompress, leaveOpen: false),
#if NET8_0_OR_GREATER
            CsvCompressionType.Brotli => new BrotliStream(stream, CompressionMode.Decompress, leaveOpen: false),
            CsvCompressionType.ZLib => new ZLibStream(stream, CompressionMode.Decompress, leaveOpen: false),
#else
            CsvCompressionType.Brotli => throw new PlatformNotSupportedException("Brotli CSV compression requires a .NET runtime that supports BrotliStream."),
            CsvCompressionType.ZLib => throw new PlatformNotSupportedException("ZLib CSV compression requires a .NET runtime that supports ZLibStream."),
#endif
            _ => throw new ArgumentOutOfRangeException(nameof(compressionType), compressionType, "Unsupported CSV compression type.")
        };

    private static Stream WrapWriteStream(Stream stream, CsvCompressionType compressionType, CompressionLevel compressionLevel) =>
        compressionType switch
        {
            CsvCompressionType.None => stream,
            CsvCompressionType.GZip => new GZipStream(stream, compressionLevel, leaveOpen: false),
            CsvCompressionType.Deflate => new DeflateStream(stream, compressionLevel, leaveOpen: false),
#if NET8_0_OR_GREATER
            CsvCompressionType.Brotli => new BrotliStream(stream, compressionLevel, leaveOpen: false),
            CsvCompressionType.ZLib => new ZLibStream(stream, compressionLevel, leaveOpen: false),
#else
            CsvCompressionType.Brotli => throw new PlatformNotSupportedException("Brotli CSV compression requires a .NET runtime that supports BrotliStream."),
            CsvCompressionType.ZLib => throw new PlatformNotSupportedException("ZLib CSV compression requires a .NET runtime that supports ZLibStream."),
#endif
            _ => throw new ArgumentOutOfRangeException(nameof(compressionType), compressionType, "Unsupported CSV compression type.")
        };

    private sealed class CsvBoundedReadStream : Stream
    {
        private readonly Stream _inner;
        private readonly long _maxBytes;
        private long _bytesRead;

        public CsvBoundedReadStream(Stream inner, long maxBytes)
        {
            _inner = inner ?? throw new ArgumentNullException(nameof(inner));
            _maxBytes = maxBytes;
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
            _bytesRead += read;
            if (_bytesRead > _maxBytes)
            {
                throw new InvalidOperationException($"CSV decompressed data exceeded the configured limit of {_maxBytes} bytes.");
            }

            return read;
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _inner.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
