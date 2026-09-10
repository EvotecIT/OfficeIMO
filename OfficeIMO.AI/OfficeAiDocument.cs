using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Reader;

namespace OfficeIMO.AI;

/// <summary>Immutable text observation projected from Reader; its identifier is scoped to the evidence snapshot hash.</summary>
public sealed record OfficeAiEvidence(string Id, string Kind, string Text, int? Page, string? SourceBlockId) {
    /// <summary>Gets the Reader anchor shared by source placeholders and their table observations, when available.</summary>
    public string? SourceAnchor { get; init; }
    /// <summary>Original source geometry where the Reader supplied it, in source coordinate units.</summary>
    public OfficeAiRegion? Region { get; init; }
}

/// <summary>Immutable source geometry; no model-generated coordinates are admitted.</summary>
public sealed record OfficeAiRegion(double X, double Y, double Width, double Height);

/// <summary>Immutable bounded source evidence. Original Reader objects and caller streams are not retained.</summary>
public sealed class OfficeAiDocument {
    private OfficeAiDocument(string hash, int sourceByteLength, OfficeAiEvidence[] evidence, int[] pages, OfficeAiImage[] images,
        string pageProvenance, bool hasSourceDiagnostics) {
        SourceHash = hash;
        SourceByteLength = sourceByteLength;
        Evidence = Array.AsReadOnly(evidence);
        Pages = Array.AsReadOnly(pages);
        Images = Array.AsReadOnly(images);
        PageProvenance = pageProvenance;
        HasSourceDiagnostics = hasSourceDiagnostics;
        SnapshotHash = Convert.ToHexString(SHA256.HashData(JsonSerializer.SerializeToUtf8Bytes(new {
            sourceHash = hash, evidence, pages, pageProvenance, hasSourceDiagnostics,
            images = images.Select(image => new { image.Id, image.Page, image.MediaType, image.Width, image.Height, image.ContentHash })
        }))).ToLowerInvariant();
    }

    /// <summary>SHA-256 of the exact captured source bytes.</summary>
    public string SourceHash { get; }
    /// <summary>SHA-256 binding source bytes to this exact evidence projection, page metadata, image payload identities and coverage state.</summary>
    public string SnapshotHash { get; }
    /// <summary>Number of original bytes captured for this snapshot.</summary>
    public int SourceByteLength { get; }
    /// <summary>Immutable text observations in source order.</summary>
    public IReadOnlyList<OfficeAiEvidence> Evidence { get; }
    /// <summary>All known one-based pages, including pages without text.</summary>
    public IReadOnlyList<int> Pages { get; }
    /// <summary>Explicit inline images supplied by the source/rendering owner.</summary>
    public IReadOnlyList<OfficeAiImage> Images { get; }
    /// <summary>Reader's native/computed/explicit/unknown page provenance.</summary>
    public string PageProvenance { get; }
    /// <summary>True when the source reader reported diagnostics; AI cannot certify complete source reconstruction.</summary>
    public bool HasSourceDiagnostics { get; }

    /// <summary>Reads and fingerprints a bounded source snapshot. The caller retains ownership of the stream.</summary>
    public static async Task<OfficeAiDocument> ReadAsync(OfficeDocumentReader reader, Stream source, string sourceName,
        OfficeAiLimits? limits = null, ReaderOptions? readerOptions = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(reader);
        ArgumentNullException.ThrowIfNull(source);
        limits ??= new(); limits.Validate();
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        deadline.CancelAfter(limits.Timeout);
        CancellationToken token = deadline.Token;
        Task<OfficeAiDocument> pending = Task.Run(async () => {
            token.ThrowIfCancellationRequested();
            byte[] bytes = await ReadBoundedAsync(source, limits.MaxInputBytes, token).ConfigureAwait(false);
            OfficeDocumentReadResult result = await reader.ReadDocumentAsync(bytes, sourceName, readerOptions, token).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            return FromReadResult(bytes, result, limits: limits);
        });
        _ = pending.ContinueWith(task => { _ = task.Exception; }, CancellationToken.None,
            TaskContinuationOptions.OnlyOnFaulted | TaskContinuationOptions.ExecuteSynchronously, TaskScheduler.Default);
        return await pending.WaitAsync(token).ConfigureAwait(false);
    }

    /// <summary>
    /// Captures a trusted adapter's read result and optional page images for the supplied original bytes.
    /// The adapter must enforce source permissions and ensure readback/images describe those bytes.
    /// This method copies observations; later mutations of the Reader result cannot change evidence.
    /// </summary>
    public static OfficeAiDocument FromReadResult(byte[] sourceBytes, OfficeDocumentReadResult document,
        IEnumerable<OfficeAiImage>? images = null, OfficeAiLimits? limits = null) {
        ArgumentNullException.ThrowIfNull(sourceBytes);
        ArgumentNullException.ThrowIfNull(document);
        limits ??= new(); limits.Validate();
        if (sourceBytes.Length is 0 || sourceBytes.Length > limits.MaxInputBytes)
            throw new ArgumentOutOfRangeException(nameof(sourceBytes), "Source exceeds the snapshot byte limit or is empty.");
        if (document.Diagnostics.Any(item => item.Severity == OfficeDocumentDiagnosticSeverity.Error))
            throw new InvalidDataException("The source reader reported an error; resolve it before model processing.");
        var evidence = new List<OfficeAiEvidence>();
        var pages = new SortedSet<int>();
        long characters = 0;
        void AddPage(int page) {
            if (page < 1) throw new InvalidDataException("Source page must be a positive one-based number.");
            pages.Add(page);
            if (pages.Count > limits.MaxPages) throw new InvalidDataException("Source page count exceeds the configured limit.");
        }
        void Add(string kind, string text, int? page, string? blockId = null, OfficeDocumentRegion? region = null, string? sourceAnchor = null) {
            if (page.HasValue) AddPage(page.Value);
            if (string.IsNullOrWhiteSpace(text)) return;
            characters += text.Length;
            if (characters > limits.MaxDocumentCharacters || evidence.Count >= limits.MaxDocumentBlocks)
                throw new InvalidDataException("Source observations exceed the configured document limits.");
            evidence.Add(new OfficeAiEvidence("e" + (evidence.Count + 1), kind, text, page, blockId) {
                SourceAnchor = sourceAnchor,
                Region = region is null ? null : new(region.X, region.Y, region.Width, region.Height)
            });
        }
        foreach (OfficeDocumentPage page in document.Pages) {
            if ((page.GetResolvedLocation().Page) is int number) AddPage(number);
        }
        var content = document.EnumerateContent().ToArray();
        var tableScopes = content.Where(item => item.Table is not null).Select(item => item.Location).ToArray();
        var pageTableCounts = tableScopes.GroupBy(location => (location?.Path ?? "", location?.Page, location?.SourceBlockIndex))
            .ToDictionary(group => group.Key, group => group.Count());
        var documentTableCounts = tableScopes.GroupBy(location => location?.Path ?? "")
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        int tableIndex = 0;
        foreach (OfficeDocumentContentItem item in content) {
            if (item.Block is { } block) {
                Add(block.Kind, block.Text, item.Location?.Page, block.Id, block.Region, item.Location?.BlockAnchor);
                continue;
            }
            if (item.Chunk is { } chunk) {
                // The PDF adapter reserves these kinds for generated notices/placeholders.
                // Other adapters use the same words for real semantic source content.
                if (chunk.Kind == ReaderInputKind.Pdf && chunk.Location?.SourceBlockKind is "warning" or "visual") continue;
                // Only the source adapter can establish that a chunk contains table text alone.
                // Retain its fallback if the matching structured table scope is incomplete.
                if (chunk.Kind == ReaderInputKind.Pdf && chunk.Location is { SourceBlockKind: "table" } location
                    && chunk.Diagnostics is { TableCount: > 0 } diagnostics) {
                    int captured = location.Page.HasValue
                        ? pageTableCounts.GetValueOrDefault((location.Path ?? "", location.Page, location.SourceBlockIndex))
                        : documentTableCounts.GetValueOrDefault(location.Path ?? "");
                    if (captured == diagnostics.TableCount) continue;
                }
                Add("chunk", chunk.Text, item.Location?.Page, chunk.Id, sourceAnchor: item.Location?.BlockAnchor);
                continue;
            }
            ReaderTable table = item.Table!;
            // Preserve table identity and empty headers; keep every row beside its table in source order.
            tableIndex++;
            if (table.Columns.Count > limits.MaxTableCells || table.Rows.Sum(row => (long)row.Count) > limits.MaxTableCells)
                throw new InvalidDataException("Source table exceeds the configured cell limit.");
            int? tablePage = item.Location?.Page;
            string? anchor = item.Location?.BlockAnchor;
            string tableId = string.IsNullOrWhiteSpace(anchor) ? $"table-{tableIndex}" : anchor;
            string title = string.IsNullOrWhiteSpace(table.Title) ? "" : table.Title + "\n";
            Add("table", title + string.Join(" | ", table.Columns), tablePage, tableId, sourceAnchor: anchor);
            int rowIndex = 0;
            foreach (IReadOnlyList<string> row in table.Rows) {
                rowIndex++;
                if (row.Count > limits.MaxTableCells) throw new InvalidDataException("Source table row exceeds the configured cell limit.");
                string text = string.Join(" | ", row.Select((value, column) =>
                    (column < table.Columns.Count ? table.Columns[column] : "Column " + (column + 1)) + ": " + value));
                Add("table-row", title + text, tablePage, $"{tableId}-row-{rowIndex}", sourceAnchor: anchor);
            }
        }
        var imageList = new List<OfficeAiImage>();
        var imageIds = new HashSet<string>(StringComparer.Ordinal);
        long totalImageBytes = 0;
        foreach (OfficeAiImage image in images ?? Array.Empty<OfficeAiImage>()) {
            ArgumentNullException.ThrowIfNull(image);
            if (imageList.Count >= limits.MaxDocumentImages || !imageIds.Add(image.Id)
                || evidence.Any(item => item.Id == image.Id)) throw new ArgumentException("Image identities or page bounds are invalid.", nameof(images));
            totalImageBytes += image.ByteLength;
            if (totalImageBytes > limits.MaxInputBytes) throw new InvalidDataException("Aggregate image evidence exceeds the snapshot byte limit.");
            imageList.Add(image); AddPage(image.Page);
        }
        bool incompleteSource = document.Diagnostics.Any(diagnostic => diagnostic.Severity != OfficeDocumentDiagnosticSeverity.Information
            || diagnostic.Category != OfficeDocumentDiagnosticCategory.Detection)
            || document.Chunks.Any(chunk => chunk.Warnings?.Count > 0)
            || document.OcrCandidates.Count > 0 || document.Pages.Any(page => page.OcrCandidates.Count > 0);
        // Table truncation is independent of top-level diagnostics. Inspect every Reader owner,
        // including chunk tables, so alternate projections cannot silently erase known omissions.
        incompleteSource |= document.Tables.Concat(document.Pages.SelectMany(page => page.Tables))
            .Concat(document.Chunks.SelectMany(chunk => chunk.Tables ?? Array.Empty<ReaderTable>()))
            .Any(table => table.Truncated || table.TotalRowCount > table.Rows.Count);
        return new OfficeAiDocument(Convert.ToHexString(SHA256.HashData(sourceBytes)).ToLowerInvariant(), sourceBytes.Length, evidence.ToArray(),
            pages.ToArray(), imageList.ToArray(), document.GetPageProvenance().ToString(), incompleteSource);
    }

    private static async Task<byte[]> ReadBoundedAsync(Stream source, int maximum, CancellationToken cancellationToken) {
        using var target = new MemoryStream();
        byte[] buffer = new byte[81920];
        while (true) {
            int read = await source.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            if (target.Length + read > maximum) throw new InvalidDataException("Source exceeds the snapshot byte limit.");
            target.Write(buffer, 0, read);
        }
        return target.ToArray();
    }
}
