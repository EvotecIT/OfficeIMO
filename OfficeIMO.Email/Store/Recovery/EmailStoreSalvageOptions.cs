using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Output, limit, and scope policy for recovering items from a damaged DBX or OLM source.</summary>
public sealed class EmailStoreSalvageOptions {
    /// <summary>Creates salvage options.</summary>
    /// <param name="format">
    /// Per-item output format. Null keeps each source's native preservation: recovered DBX MIME bytes are written
    /// unchanged as EML, and OLM items are written as MSG.
    /// </param>
    /// <param name="overwriteExisting">Whether existing artifacts and the manifest may be replaced.</param>
    /// <param name="writeManifest">Whether a provenance manifest is written next to the recovered items.</param>
    /// <param name="maxRecoveredItems">Maximum item outcomes produced before recovery stops.</param>
    /// <param name="includeUnindexedFragmentChains">
    /// Whether DBX message fragment chains that no surviving index record references are recovered when their first
    /// fragment starts with an RFC 822 header.
    /// </param>
    /// <param name="readerOptions">Source bounds. The store reader defaults apply when null.</param>
    /// <param name="writerOptions">
    /// Item writer bounds and conversion policy. When null, conversion loss is reported in each item's diagnostics
    /// instead of rejecting the item, because recovery favours getting readable data out of a damaged source.
    /// </param>
    public EmailStoreSalvageOptions(
        EmailFileFormat? format = null,
        bool overwriteExisting = false,
        bool writeManifest = true,
        int maxRecoveredItems = 10_000,
        bool includeUnindexedFragmentChains = true,
        EmailStoreReaderOptions? readerOptions = null,
        EmailWriterOptions? writerOptions = null) {
        if (format.HasValue && format.Value != EmailFileFormat.Eml && format.Value != EmailFileFormat.OutlookMsg &&
            format.Value != EmailFileFormat.OutlookTemplate && format.Value != EmailFileFormat.Tnef) {
            throw new ArgumentException("Salvage supports EML, MSG, OFT, or TNEF output.", nameof(format));
        }
        ArgumentOutOfRangeException.ThrowIfNegativeOrZero(maxRecoveredItems);
        Format = format;
        OverwriteExisting = overwriteExisting;
        WriteManifest = writeManifest;
        MaxRecoveredItems = maxRecoveredItems;
        IncludeUnindexedFragmentChains = includeUnindexedFragmentChains;
        ReaderOptions = readerOptions ?? EmailStoreReaderOptions.Default;
        WriterOptions = writerOptions ?? new EmailWriterOptions(EmailConversionLossPolicy.Warn);
    }

    /// <summary>Requested output format, or null for the source's native preservation format.</summary>
    public EmailFileFormat? Format { get; }
    /// <summary>Whether existing artifacts and the manifest may be replaced.</summary>
    public bool OverwriteExisting { get; }
    /// <summary>Whether a provenance manifest is written.</summary>
    public bool WriteManifest { get; }
    /// <summary>Maximum item outcomes produced before recovery stops.</summary>
    public int MaxRecoveredItems { get; }
    /// <summary>Whether unreferenced DBX fragment chains that start with a message header are recovered.</summary>
    public bool IncludeUnindexedFragmentChains { get; }
    /// <summary>Source bounds applied while scanning, decoding, and projecting recovered data.</summary>
    public EmailStoreReaderOptions ReaderOptions { get; }
    /// <summary>Item writer bounds and conversion policy; conversion loss is reported, not rejected, by default.</summary>
    public EmailWriterOptions WriterOptions { get; }
}
