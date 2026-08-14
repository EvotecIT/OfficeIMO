#nullable enable

using System.Data;
using System.Threading;

namespace OfficeIMO.CSV;

public sealed partial class CsvDocument
{
    /// <summary>
    /// Writes an open data reader to CSV with bounded, ordered parallel row formatting.
    /// </summary>
    /// <param name="writer">Destination text writer.</param>
    /// <param name="reader">Source data reader positioned before the first row.</param>
    /// <param name="options">Optional CSV serialization settings.</param>
    /// <param name="parallelOptions">Optional worker and batch limits.</param>
    /// <param name="cancellationToken">Token observed while reading, formatting, and writing rows.</param>
    public static void WriteDataReaderParallel(
        TextWriter writer,
        IDataReader reader,
        CsvSaveOptions? options = null,
        CsvWriteParallelOptions? parallelOptions = null,
        CancellationToken cancellationToken = default)
    {
        if (writer == null) throw new ArgumentNullException(nameof(writer));
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        options ??= new CsvSaveOptions();
        if (options.Append || options.NoClobber)
        {
            throw new ArgumentException("Append and NoClobber apply only to path writes.", nameof(options));
        }

        WriteDataReaderParallelCore(writer, reader, options, parallelOptions, cancellationToken);
    }

    /// <summary>
    /// Writes an open data reader to a CSV file with bounded, ordered parallel row formatting.
    /// </summary>
    /// <param name="path">Destination CSV path.</param>
    /// <param name="reader">Source data reader positioned before the first row.</param>
    /// <param name="options">Optional CSV serialization settings.</param>
    /// <param name="parallelOptions">Optional worker and batch limits.</param>
    /// <param name="cancellationToken">Token observed while reading, formatting, and writing rows.</param>
    public static void WriteDataReaderParallel(
        string path,
        IDataReader reader,
        CsvSaveOptions? options = null,
        CsvWriteParallelOptions? parallelOptions = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        options ??= new CsvSaveOptions();
        cancellationToken.ThrowIfCancellationRequested();
        WritePath(
            path,
            options,
            writer => WriteDataReaderParallelCore(writer, reader, options, parallelOptions, cancellationToken));
    }

    /// <summary>
    /// Writes an open data reader to a CSV stream with bounded, ordered parallel row formatting.
    /// </summary>
    /// <param name="destination">Writable destination stream.</param>
    /// <param name="reader">Source data reader positioned before the first row.</param>
    /// <param name="options">Optional CSV serialization settings.</param>
    /// <param name="parallelOptions">Optional worker and batch limits.</param>
    /// <param name="leaveOpen">Whether the destination stream remains open after writing.</param>
    /// <param name="cancellationToken">Token observed while reading, formatting, and writing rows.</param>
    public static void WriteDataReaderParallel(
        Stream destination,
        IDataReader reader,
        CsvSaveOptions? options = null,
        CsvWriteParallelOptions? parallelOptions = null,
        bool leaveOpen = true,
        CancellationToken cancellationToken = default)
    {
        if (destination == null) throw new ArgumentNullException(nameof(destination));
        if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        options ??= new CsvSaveOptions();
        if (options.Append || options.NoClobber)
        {
            throw new ArgumentException("Append and NoClobber apply only to path writes.", nameof(options));
        }

        using var writer = CsvFile.CreateTextWriter(destination, options, leaveOpen, FileBufferSize);
        WriteDataReaderParallelCore(writer, reader, options, parallelOptions, cancellationToken);
    }

    private static void WriteDataReaderParallelCore(
        TextWriter writer,
        IDataReader reader,
        CsvSaveOptions options,
        CsvWriteParallelOptions? parallelOptions,
        CancellationToken cancellationToken)
    {
        using var rowWriter = new CsvRowWriter(writer, options, leaveOpen: true);
        rowWriter.WriteDataReaderParallel(reader, parallelOptions, cancellationToken);
    }
}
