#nullable enable

namespace OfficeIMO.CSV;

public sealed partial class CsvRowWriter
{
    /// <summary>Creates a row writer at the destination stream's current position.</summary>
    /// <param name="destination">Writable destination stream.</param>
    /// <param name="options">CSV serialization, encoding, and compression settings. Auto compression means none for streams.</param>
    /// <param name="leaveOpen">Whether disposal keeps the supplied stream open. Compression wrappers are always completed and disposed.</param>
    /// <param name="bufferSize">Text and byte buffer size.</param>
    /// <returns>A writer that owns its buffering and observes the requested stream ownership.</returns>
    public static CsvRowWriter CreateStream(
        Stream destination,
        CsvSaveOptions? options = null,
        bool leaveOpen = false,
        int bufferSize = 256 * 1024)
    {
        options ??= new CsvSaveOptions();
        if (bufferSize <= 0) throw new ArgumentOutOfRangeException(nameof(bufferSize));
        return new CsvRowWriter(CsvFile.CreateTextWriter(destination, options, leaveOpen, bufferSize), options);
    }
}
