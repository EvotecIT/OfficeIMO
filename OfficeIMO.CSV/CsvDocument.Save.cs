#nullable enable

using OfficeIMO.Core.Internal;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>
    /// Saves the document to the specified path.
    /// </summary>
    public void Save(string path, CsvSaveOptions? options = null)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("File path cannot be empty.", nameof(path));
        }

        options = ResolveSaveOptions(options);

        WritePath(path, options, writer => CsvWriter.Write(writer, this, options));
    }

    private static void WritePath(string path, CsvSaveOptions options, Action<TextWriter> write)
    {
        string fullPath = Path.GetFullPath(path);
        if (options.NoClobber && File.Exists(fullPath))
        {
            throw new IOException($"The file '{fullPath}' already exists.");
        }

        CsvCompressionType compressionType = CsvFile.ResolveCompression(options.CompressionType, fullPath);
        if (options.Append && compressionType != CsvCompressionType.None)
        {
            throw new NotSupportedException("Appending to compressed CSV files is not supported.");
        }

        OfficeFileCommit.EnsureTargetDirectory(fullPath);
        if (options.Append)
        {
            using var appendWriter = CsvFile.CreateTextWriter(fullPath, options, append: true, bufferSize: FileBufferSize);
            write(appendWriter);
            return;
        }

        string temporaryPath = OfficeFileCommit.CreateTemporaryPath(fullPath);
        try
        {
            using (var writer = CsvFile.CreateTextWriterForCompressionPath(
                       temporaryPath, fullPath, options, FileBufferSize))
            {
                write(writer);
            }

            OfficeFileCommit.CommitTemporaryFile(
                temporaryPath,
                fullPath,
                options.NoClobber
                    ? OfficeFileCommit.ConflictPolicy.FailIfExists
                    : OfficeFileCommit.ConflictPolicy.Replace);
            temporaryPath = string.Empty;
        }
        finally
        {
            OfficeFileCommit.DeleteIfExists(temporaryPath);
        }
    }

    /// <summary>Saves the document to a caller-owned writable stream.</summary>
    /// <remarks>
    /// Serialization completes before the destination is written. Successful saves replace and rewind
    /// seekable streams; non-seekable streams receive output at their current position. The destination stays open.
    /// </remarks>
    public void Save(Stream destination, CsvSaveOptions? options = null)
    {
        using var serialized = SerializeToMemoryStream(options);
        OfficeStreamWriter.WriteAllBytes(destination,
            new ArraySegment<byte>(serialized.GetBuffer(), 0, checked((int)serialized.Length)));
    }

    /// <summary>Asynchronously saves the document to a path.</summary>
    public async Task SaveAsync(string path, CsvSaveOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        options = ResolveSaveOptions(options);
        cancellationToken.ThrowIfCancellationRequested();
        string fullPath = Path.GetFullPath(path);
        if (options.NoClobber && File.Exists(fullPath)) throw new IOException($"The file '{fullPath}' already exists.");
        CsvCompressionType compressionType = CsvFile.ResolveCompression(options.CompressionType, fullPath);

        if (options.Append)
        {
            if (compressionType != CsvCompressionType.None)
                throw new NotSupportedException("Appending to compressed CSV files is not supported.");
            OfficeFileCommit.EnsureTargetDirectory(fullPath);
            using var stream = new FileStream(fullPath, options.NoClobber ? FileMode.CreateNew : FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read,
                FileBufferSize, FileOptions.Asynchronous);
            Encoding appendEncoding = options.Encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
            bool needsSeparator = CsvFile.NeedsAppendRecordSeparator(stream, appendEncoding);
            stream.Position = stream.Length;
            await SaveToStreamAsync(stream, CopySaveOptions(options, CsvCompressionType.None), cancellationToken,
                needsSeparator ? options.NewLine : null).ConfigureAwait(false);
            return;
        }

        await OfficeFileCommit.WriteAsync(fullPath,
            (stream, token) => SaveToStreamAsync(stream, CopySaveOptions(options, compressionType), token),
            options.NoClobber ? OfficeFileCommit.ConflictPolicy.FailIfExists : OfficeFileCommit.ConflictPolicy.Replace,
            cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously saves the document to a caller-owned writable stream.</summary>
    /// <remarks>
    /// Successful saves replace and rewind seekable streams. Non-seekable streams receive output at their
    /// current position. The destination stays open and can contain partial output if saving fails.
    /// </remarks>
    public Task SaveAsync(Stream destination, CsvSaveOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        cancellationToken.ThrowIfCancellationRequested();
        options = ResolveSaveOptions(options);
        if (options.Append || options.NoClobber)
            throw new ArgumentException("Append and NoClobber apply only to path saves.", nameof(options));
        return OfficeStreamWriter.WriteAsync(destination,
            (stream, token) => SaveToStreamAsync(stream, options, token), cancellationToken);
    }

    private async Task SaveToStreamAsync(Stream destination, CsvSaveOptions options, CancellationToken cancellationToken, string? initialSeparator = null)
    {
        using var guard = new CsvFile.AsyncWriteGuard(destination, cancellationToken);
        var writer = CsvFile.CreateTextWriter(guard, options, leaveOpen: true, FileBufferSize);
        try {
            await CsvWriter.WriteAsync(writer, this, options, cancellationToken, initialSeparator, guard.DrainAsync).ConfigureAwait(false);
#if NET8_0_OR_GREATER
            await writer.DisposeAsync().ConfigureAwait(false);
#else
            writer.Dispose();
#endif
            await guard.DrainAsync().ConfigureAwait(false);
            await guard.FlushAsync(cancellationToken).ConfigureAwait(false);
        } catch {
            guard.SuppressWrites = true;
            writer.Dispose();
            throw;
        }
    }

    /// <summary>Encodes the document using the selected CSV encoding and compression.</summary>
    public byte[] ToBytes(CsvSaveOptions? options = null)
    {
        using var stream = SerializeToMemoryStream(options);
        return stream.ToArray();
    }

    /// <summary>Encodes the document in a new writable memory stream positioned at the beginning.</summary>
    public MemoryStream ToStream(CsvSaveOptions? options = null) => new MemoryStream(ToBytes(options));

    // Stage the complete artifact so serialization failures cannot overwrite a caller's existing document.
    private MemoryStream SerializeToMemoryStream(CsvSaveOptions? options)
    {
        options = ResolveSaveOptions(options);
        if (options.Append || options.NoClobber)
            throw new ArgumentException("Append and NoClobber apply only to path saves.", nameof(options));
        CsvCompressionType compressionType = options.CompressionType == CsvCompressionType.Auto
            ? CsvCompressionType.None
            : options.CompressionType;
        var serializationOptions = CopySaveOptions(options, compressionType);
        var stream = new MemoryStream();
        try
        {
            using (TextWriter writer = CsvFile.CreateTextWriter(stream, serializationOptions, leaveOpen: true, FileBufferSize))
            {
                CsvWriter.Write(writer, this, serializationOptions);
            }
            return stream;
        }
        catch
        {
            stream.Dispose();
            throw;
        }
    }

    private static CsvSaveOptions CopySaveOptions(CsvSaveOptions source, CsvCompressionType compressionType)
    {
        return new CsvSaveOptions
        {
            Delimiter = source.Delimiter,
            DelimiterText = source.DelimiterText,
            NewLine = source.NewLine,
            IncludeHeader = source.IncludeHeader,
            Culture = source.Culture,
            Encoding = source.Encoding,
            CompressionType = compressionType,
            CompressionLevel = source.CompressionLevel,
            NullValue = source.NullValue,
            DateTimeFormat = source.DateTimeFormat,
            UseUtc = source.UseUtc,
            FormulaInjectionPolicy = source.FormulaInjectionPolicy,
            QuoteMode = source.QuoteMode,
            QuoteFields = source.QuoteFields
        };
    }

    private CsvSaveOptions ResolveSaveOptions(CsvSaveOptions? options)
    {
        return options ?? new CsvSaveOptions
        {
            Delimiter = _delimiter,
            DelimiterText = _delimiterText,
            Culture = _culture,
            Encoding = _encoding
        };
    }

}
