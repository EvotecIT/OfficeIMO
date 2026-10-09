using OfficeIMO.DjVu;
using OfficeIMO.Drawing;

namespace OfficeIMO.Reader.DjVu;

internal static class DjVuReaderAdapter {
    internal static OfficeDocumentReadResult Read(string path, ReaderOptions reader, ReaderDjVuOptions options, CancellationToken token) {
        var limits = Limits(reader, options);
        var input = DocumentReaderEngine.ReadAdapterInput(path, reader, token, limits.MaxSourceBytes);
        return ReadSnapshot(input, reader, options, limits, token);
    }
    internal static OfficeDocumentReadResult Read(Stream stream, string? name, ReaderOptions reader, ReaderDjVuOptions options, CancellationToken token) {
        var limits = Limits(reader, options);
        var input = DocumentReaderEngine.ReadAdapterInput(stream, name ?? "document.djvu", reader, token, limits.MaxSourceBytes);
        return ReadSnapshot(input, reader, options, limits, token);
    }
    private static OfficeDocumentReadResult ReadSnapshot(ReaderAdapterInputSnapshot input, ReaderOptions reader, ReaderDjVuOptions options, DjVuReadOptions limits, CancellationToken token) {
        var result = Project(DjVuDocument.Load(input.Bytes, limits, token), input.Source, reader, options, token);
        foreach (var chunk in result.Chunks) DocumentReaderEngine.ApplyAdapterSource(chunk, input, reader.ComputeHashes);
        return result;
    }
    private static DjVuReadOptions Limits(ReaderOptions reader, ReaderDjVuOptions options) {
        var limits = options.ReadOptions.Clone();
        if (reader.MaxInputBytes.HasValue) limits.MaxSourceBytes = Math.Min(limits.MaxSourceBytes, reader.MaxInputBytes.Value);
        return limits;
    }

    internal static OfficeDocumentReadResult Project(DjVuDocument document, OfficeDocumentSource source, ReaderOptions reader, ReaderDjVuOptions options, CancellationToken token) {
        var limits = reader.ResourceLimits?.CloneValidated();
        int maxChars = Math.Max(1, reader.MaxChars);
        var chunks = new List<ReaderChunk>();
        var blocks = new List<OfficeDocumentBlock>();
        var pages = new List<OfficeDocumentPage>();
        var assets = new List<OfficeDocumentAsset>();
        var candidates = new List<OfficeDocumentOcrCandidate>();
        var diagnostics = new List<OfficeDocumentDiagnostic>();
        var metadata = new List<OfficeDocumentMetadataEntry>();
        var links = new List<OfficeDocumentLink>();
        metadata.Add(new OfficeDocumentMetadataEntry { Id = "djvu-page-count", Category = "reader.container", Name = "PageCount",
            Value = document.Pages.Count.ToString(CultureInfo.InvariantCulture), ValueType = "count" });
        long assetBytes = 0, chunkCharacters = 0;
        foreach (var page in DjVuPageSelection.Select(document, options.PageNumbers, options.ReadOptions.MaxPages, token)) {
            token.ThrowIfCancellationRequested();
            var location = new ReaderLocation { Path = source.Path, Page = page.Number, LogicalOrder = page.Number };
            var text = page.GetText(token);
            var pageBlocks = new List<OfficeDocumentBlock>();
            int part = 0;
            foreach (string value in DocumentReaderEngine.EnumerateAdapterProjection(text.Text, maxChars, maxChars)) {
                token.ThrowIfCancellationRequested();
                Check(chunks.Count + 1L, limits?.MaxChunks, nameof(ReaderResourceLimits.MaxChunks));
                Check(chunkCharacters + value.Length * 2L, limits?.MaxChunkCharacters, nameof(ReaderResourceLimits.MaxChunkCharacters));
                chunkCharacters += value.Length * 2L;
                var chunkLocation = new ReaderLocation { Path = source.Path, Page = page.Number, LogicalOrder = chunks.Count + 1, BlockAnchor = "djvu-page-" + page.Number + "-part-" + part };
                chunks.Add(new ReaderChunk { Id = chunkLocation.BlockAnchor!, Kind = ReaderInputKind.DjVu, Location = chunkLocation,
                    Text = value, Markdown = value, ContinuesPreviousChunk = part++ > 0,
                    SourceId = source.SourceId, SourceHash = source.SourceHash, SourceLengthBytes = source.LengthBytes,
                    SourceLastWriteUtc = source.LastWriteUtc, TokenEstimate = Math.Max(1, (value.Length + 3) / 4) });
            }
            foreach (var zone in WordZones(text.Zones)) {
                token.ThrowIfCancellationRequested();
                Check(blocks.Count + 1L, limits?.MaxBlocks, nameof(ReaderResourceLimits.MaxBlocks));
                var display = page.GetDisplayBounds(zone.Bounds);
                var block = new OfficeDocumentBlock { Id = "djvu-page-" + page.Number + "-word-" + pageBlocks.Count, Kind = "word",
                    Text = text.Text.Substring(zone.CharacterOffset, zone.CharacterLength),
                    Location = new ReaderLocation { Path = source.Path, Page = page.Number, SourceBlockKind = "stored-text", SourceBlockIndex = pageBlocks.Count },
                    Region = Points(display, page.Dpi) };
                blocks.Add(block); pageBlocks.Add(block);
            }
            if (pageBlocks.Count == 0 && text.Text.Length != 0) {
                Check(blocks.Count + 1L, limits?.MaxBlocks, nameof(ReaderResourceLimits.MaxBlocks));
                var block = new OfficeDocumentBlock { Id = "djvu-page-" + page.Number + "-text", Kind = "paragraph", Text = text.Text, Location = location };
                blocks.Add(block); pageBlocks.Add(block);
            }
            bool missing = text.Status == DjVuTextStatus.Absent || text.Status == DjVuTextStatus.Empty;
            metadata.Add(new OfficeDocumentMetadataEntry { Id = "djvu-page-" + page.Number + "-text-status", Category = "djvu", Name = "stored-text-status", Value = text.Status.ToString(), ValueType = "string", SourceObjectId = page.Id, Location = location });
            metadata.Add(new OfficeDocumentMetadataEntry { Id = "djvu-page-" + page.Number + "-rotation", Category = "djvu", Name = "source-rotation-degrees", Value = page.Rotation.ToString(CultureInfo.InvariantCulture), ValueType = "number", SourceObjectId = page.Id, Location = location });
            if (text.Status == DjVuTextStatus.Corrupt) diagnostics.Add(new OfficeDocumentDiagnostic {
                Code = "djvu.text.corrupt", Message = text.Diagnostic ?? "Stored DjVu text is corrupt.", Severity = OfficeDocumentDiagnosticSeverity.Warning,
                Category = OfficeDocumentDiagnosticCategory.Parsing, Source = "OfficeIMO.DjVu", IsRecoverable = true, Location = location });
            OfficeDocumentAsset? asset = null;
            if (options.ImageMode == ReaderDjVuImageMode.AllPages || options.ImageMode == ReaderDjVuImageMode.MissingTextPages && missing) {
                Check(assets.Count + 1L, options.MaxPageImages, nameof(ReaderDjVuOptions.MaxPageImages));
                Check(assets.Count + 1L, limits?.MaxAssets, nameof(ReaderResourceLimits.MaxAssets));
                long remaining = Math.Min(options.MaxTotalPageImageBytes - assetBytes, (limits?.MaxAssetBytes ?? long.MaxValue) - assetBytes);
                if (remaining <= 0) throw new ReaderResourceLimitException(nameof(ReaderResourceLimits.MaxAssetBytes), limits?.MaxAssetBytes ?? options.MaxTotalPageImageBytes);
                var rendered = page.Render(options.RenderOptions, token);
                byte[] payload = OfficeRasterImageEncoder.Encode(rendered.Image, OfficeImageExportFormat.Png, null,
                    Math.Min(options.MaxPageImageBytes, remaining), token);
                assetBytes += payload.LongLength;
                asset = new OfficeDocumentAsset { Id = "djvu-page-" + page.Number + "-image", Kind = "image", MediaType = "image/png", Extension = ".png",
                    FileName = "page-" + page.Number.ToString(CultureInfo.InvariantCulture) + ".png", SourceObjectId = page.Id,
                    Width = rendered.Image.Width, Height = rendered.Image.Height, PayloadBytes = payload, LengthBytes = payload.LongLength,
                    PayloadHash = OfficeDocumentAssetHash.ComputeSha256Hex(payload), Location = location,
                    Region = new OfficeDocumentRegion { Width = page.DisplayWidth * 72.0 / page.Dpi, Height = page.DisplayHeight * 72.0 / page.Dpi } };
                assets.Add(asset);
                foreach (var diagnostic in rendered.FidelityDiagnostics) diagnostics.Add(new OfficeDocumentDiagnostic { Code = diagnostic.Code, Message = diagnostic.Message,
                    Severity = OfficeDocumentDiagnosticSeverity.Warning, Category = OfficeDocumentDiagnosticCategory.Content, Source = "OfficeIMO.DjVu", IsRecoverable = true, Location = location });
            }
            OfficeDocumentOcrCandidate? candidate = null;
            if (missing) {
                candidate = new OfficeDocumentOcrCandidate { Id = "djvu-page-" + page.Number + "-ocr", Kind = "page", AssetId = asset?.Id,
                    Reason = text.Status == DjVuTextStatus.Absent ? "The page has no stored text layer." : "The stored text layer is empty.",
                    ImageCount = 1, TextBlockCount = 0, Location = location,
                    Region = new OfficeDocumentRegion { Width = page.DisplayWidth * 72.0 / page.Dpi, Height = page.DisplayHeight * 72.0 / page.Dpi } };
                candidates.Add(candidate);
            }
            pages.Add(new OfficeDocumentPage { Number = page.Number, Name = page.Title, Location = location,
                Width = page.DisplayWidth * 72.0 / page.Dpi, Height = page.DisplayHeight * 72.0 / page.Dpi, RotationDegrees = 0,
                Blocks = pageBlocks.AsReadOnly(), Assets = asset == null ? Array.Empty<OfficeDocumentAsset>() : new[] { asset },
                OcrCandidates = candidate == null ? Array.Empty<OfficeDocumentOcrCandidate>() : new[] { candidate } });
        }
        foreach (var bookmark in Bookmarks(document.Bookmarks)) {
            token.ThrowIfCancellationRequested();
            links.Add(new OfficeDocumentLink { Id = "djvu-bookmark-" + links.Count, Text = bookmark.Title,
                Kind = bookmark.PageNumber.HasValue ? "destination" : "uri", Uri = bookmark.Target.Length == 0 ? null : bookmark.Target,
                DestinationPageNumber = bookmark.PageNumber, Location = new ReaderLocation { Path = source.Path, SourceBlockKind = "outline" } });
        }
        if (document.BookmarkDiagnostic != null) diagnostics.Add(new OfficeDocumentDiagnostic { Code = "djvu.outline.corrupt",
            Message = document.BookmarkDiagnostic, Severity = OfficeDocumentDiagnosticSeverity.Warning,
            Category = OfficeDocumentDiagnosticCategory.Parsing, Source = "OfficeIMO.DjVu", IsRecoverable = true });
        var result = DocumentReaderEngine.CreateRichDocumentResult(chunks, ReaderInputKind.DjVu, source,
            new[] { "officeimo.djvu.stored-text", "officeimo.djvu.pages", "officeimo.djvu.display-geometry" }, blocks, Array.Empty<ReaderTable>(), pages, assets, links);
        result.OcrCandidates = candidates.AsReadOnly(); result.Diagnostics = diagnostics.AsReadOnly(); result.Metadata = result.Metadata.Concat(metadata).ToArray();
        if (reader.ComputeHashes) foreach (var chunk in chunks) chunk.ChunkHash = DocumentReaderEngine.ComputeChunkHash(chunk);
        if (chunks.Count > 0) chunks[0].Warnings = diagnostics.Select(d => d.Message).ToArray();
        return result;
    }
    private static IEnumerable<DjVuTextZone> WordZones(IEnumerable<DjVuTextZone> zones) {
        foreach (var zone in zones) {
            if (zone.Kind == DjVuTextZoneKind.Word) yield return zone;
            else foreach (var word in WordZones(zone.Children)) yield return word;
        }
    }
    private static IEnumerable<DjVuBookmark> Bookmarks(IEnumerable<DjVuBookmark> bookmarks) {
        foreach (var bookmark in bookmarks) {
            yield return bookmark;
            foreach (var child in Bookmarks(bookmark.Children)) yield return child;
        }
    }
    private static OfficeDocumentRegion Points(DjVuRectangle bounds, int dpi) => new OfficeDocumentRegion {
        X = bounds.X * 72.0 / dpi, Y = bounds.Y * 72.0 / dpi, Width = bounds.Width * 72.0 / dpi, Height = bounds.Height * 72.0 / dpi };
    private static void Check(long value, long? maximum, string name) {
        if (maximum.HasValue && value > maximum.Value) throw new ReaderResourceLimitException(name, maximum.Value);
    }
}
