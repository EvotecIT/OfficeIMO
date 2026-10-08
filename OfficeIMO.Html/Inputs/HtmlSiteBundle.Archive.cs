using System.IO.Compression;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Html;

public sealed partial class HtmlSiteBundle {
    private static async Task<HtmlSiteBundle> ReadArchiveAsync(byte[] bytes, HtmlSiteBundleOptions options,
        HtmlConversionDocumentOptions? htmlOptions, bool asynchronous, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeArchiveSafety.ZipCentralDirectoryScanResult scan = OfficeArchiveSafety.ScanZipCentralDirectory(bytes, options.MaximumEntryCount);
        if (!scan.IsValid) throw new InvalidDataException(scan.Error ?? "The ZIP site bundle is malformed.");
        if (scan.LimitExceeded) throw new InvalidDataException("The ZIP site bundle exceeds its entry-count limit.");

        using var source = new MemoryStream(bytes, writable: false);
        using var archive = new ZipArchive(source, ZipArchiveMode.Read, leaveOpen: true);
        var metadata = OfficeIMO.Provenance.OfficeProvenanceZip.GetInputEntryMetadata(bytes, archive);
        var entries = new Dictionary<string, ZipArchiveEntry>(StringComparer.Ordinal);
        var allNames = new HashSet<string>(StringComparer.Ordinal);
        long decodedBytes = 0;
        foreach (ZipArchiveEntry entry in archive.Entries) {
            uint checksum = metadata[entry].Checksum;
            cancellationToken.ThrowIfCancellationRequested();
            string name = OfficeArchiveSafety.NormalizeEntryName(metadata[entry].Name);
            if (OfficeArchiveSafety.IsUnsafePath(name)) throw new InvalidDataException("The ZIP site bundle contains an unsafe entry path.");
            if (!allNames.Add(name)) throw new InvalidDataException("The ZIP site bundle contains duplicate normalized entry paths.");
            bool directory = name.EndsWith("/", StringComparison.Ordinal);
            int unixType = (int)((metadata[entry].ExternalAttributes >> 16) & 0xf000U);
            if (unixType != 0 && unixType != 0x8000 && unixType != 0x4000) {
                throw new InvalidDataException("The ZIP site bundle contains a non-regular entry.");
            }
            long length = entry.Length;
            if (length < 0 || length > options.MaximumEntryBytes
                || length > options.MaximumTotalDecodedBytes - decodedBytes) {
                throw new InvalidDataException("The ZIP site bundle exceeds its decoded entry or total byte limit.");
            }
            if (length > 0 && (entry.CompressedLength <= 0
                || OfficeArchiveSafety.IsCompressionRatioExceeded(entry, length, options.MaximumCompressionRatio))) {
                throw new InvalidDataException("The ZIP site bundle exceeds its compression-ratio limit.");
            }
            if (directory) {
                if (length != 0 || checksum != 0) throw new InvalidDataException("ZIP directory entries cannot contain payload bytes.");
                using Stream directoryPayload = OfficeArchiveSafety.OpenEntryPayload(bytes, entry, metadata[entry]);
                _ = OfficeArchiveSafety.ReadEntryBytes(directoryPayload, 0, 0, cancellationToken);
                continue;
            }
            if (unixType == 0x4000) throw new InvalidDataException("ZIP directory metadata conflicts with its entry path.");
            decodedBytes += length;
            entries.Add(name, entry);
        }

        string entryPath = SelectEntry(entries, options.EntryPath);
        Uri baseUri = CreateEntryUri(options.ArchiveBaseUri, entryPath);
        HtmlConversionDocumentOptions parsing = (htmlOptions ?? new HtmlConversionDocumentOptions()).Clone();
        if (parsing.BaseUri != null && parsing.BaseUri != baseUri) {
            throw new ArgumentException("Use ArchiveBaseUri to select the bundle's virtual root; HTML BaseUri must match the selected entry.", nameof(htmlOptions));
        }
        parsing.BaseUri = baseUri;
        var resources = new Dictionary<string, HtmlResolvedResource>(StringComparer.Ordinal);
        byte[]? htmlBytes = null;
        foreach (KeyValuePair<string, ZipArchiveEntry> item in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            using Stream payload = OfficeArchiveSafety.OpenEntryPayload(bytes, item.Value, metadata[item.Value]);
            byte[] data = asynchronous
                ? await OfficeArchiveSafety.ReadEntryBytesAsync(payload, item.Value.Length, options.MaximumEntryBytes, cancellationToken).ConfigureAwait(false)
                : OfficeArchiveSafety.ReadEntryBytes(payload, item.Value.Length, options.MaximumEntryBytes, cancellationToken);
            OfficeArchiveSafety.ValidateEntryChecksum(data, metadata[item.Value].Checksum, cancellationToken);
            if (item.Key == entryPath) htmlBytes = data;
            string key = ResourceKey(CreateEntryUri(options.ArchiveBaseUri, item.Key));
            if (resources.ContainsKey(key)) throw new InvalidDataException("The ZIP site bundle contains colliding resource URI identities.");
            resources.Add(key, new HtmlResolvedResource(data, ContentType(item.Key)));
        }
        using var htmlStream = new MemoryStream(htmlBytes!, writable: false);
        // The async HTML loader also passes cancellation into parsing; the synchronous
        // entry point waits on this memory-backed operation without a captured context.
        HtmlConversionDocument document = await HtmlConversionDocument.LoadAsync(htmlStream, parsing,
            cancellationToken: cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return new HtmlSiteBundle(document, entryPath, options.ArchiveBaseUri, resources,
            entries.Keys.OrderBy(name => name, StringComparer.Ordinal).ToList().AsReadOnly(), bytes.LongLength, decodedBytes);
    }

    private static string SelectEntry(Dictionary<string, ZipArchiveEntry> entries, string? requested) {
        if (requested != null) {
            string normalized = OfficeArchiveSafety.NormalizeEntryName(requested);
            if (OfficeArchiveSafety.IsUnsafePath(normalized) || !IsHtml(normalized)
                || !entries.ContainsKey(normalized)) {
                throw new InvalidDataException("The requested HTML entry is absent or invalid.");
            }
            return normalized;
        }
        if (entries.ContainsKey("index.html")) return "index.html";
        if (entries.ContainsKey("index.htm")) return "index.htm";
        string[] candidates = entries.Keys.Where(IsHtml).Take(2).ToArray();
        if (candidates.Length != 1) throw new InvalidDataException("The ZIP site bundle must contain one HTML entry or specify EntryPath.");
        return candidates[0];
    }

    private static bool IsHtml(string name) => name.EndsWith(".html", StringComparison.OrdinalIgnoreCase)
        || name.EndsWith(".htm", StringComparison.OrdinalIgnoreCase);

    private static string ContentType(string name) => Path.GetExtension(name).ToLowerInvariant() switch {
        ".html" or ".htm" => "text/html",
        ".css" => "text/css",
        ".png" => "image/png",
        ".jpg" or ".jpeg" => "image/jpeg",
        ".gif" => "image/gif",
        ".webp" => "image/webp",
        ".avif" => "image/avif",
        ".svg" => "image/svg+xml",
        ".bmp" => "image/bmp",
        ".tif" or ".tiff" => "image/tiff",
        ".woff" => "font/woff",
        ".woff2" => "font/woff2",
        ".ttf" or ".ttc" => "font/ttf",
        ".otf" => "font/otf",
        _ => "application/octet-stream"
    };
}
