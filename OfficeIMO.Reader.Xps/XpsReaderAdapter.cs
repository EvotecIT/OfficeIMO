using OfficeIMO.Xps;

namespace OfficeIMO.Reader.Xps;

internal static partial class XpsReaderAdapter {
    internal static OfficeDocumentReadResult Read(string path, ReaderOptions reader, ReaderXpsOptions options, CancellationToken token) {
        var limits = Limits(reader, options);
        var input = DocumentReaderEngine.ReadAdapterInput(path, reader, token, limits.MaximumInputBytes);
        var result = Project(XpsDocument.Load(input.Bytes, limits, token), input.Source, reader, options, token);
        foreach (var chunk in result.Chunks) DocumentReaderEngine.ApplyAdapterSource(chunk, input, reader.ComputeHashes);
        return result;
    }
    internal static OfficeDocumentReadResult Read(Stream stream, string? sourceName, ReaderOptions reader, ReaderXpsOptions options, CancellationToken token) {
        var limits = Limits(reader, options);
        var input = DocumentReaderEngine.ReadAdapterInput(stream, sourceName ?? "document.xps", reader, token, limits.MaximumInputBytes);
        var result = Project(XpsDocument.Load(input.Bytes, limits, token), input.Source, reader, options, token);
        foreach (var chunk in result.Chunks) DocumentReaderEngine.ApplyAdapterSource(chunk, input, reader.ComputeHashes);
        return result;
    }
    private static XpsReadOptions Limits(ReaderOptions reader, ReaderXpsOptions options) {
        var limits = options.ReadOptions.Clone();
        if (reader.MaxInputBytes is long maximum) limits.MaximumInputBytes = (int)Math.Min(limits.MaximumInputBytes, maximum);
        return limits;
    }

    internal static OfficeDocumentReadResult Project(XpsDocument document, OfficeDocumentSource source, ReaderOptions reader,
        ReaderXpsOptions options, CancellationToken token) {
        var model = document.ToOfficeDocumentModel(source.Path, options.IncludeSvgPreviewAssets, token);
        var chunks = new List<ReaderChunk>();
        foreach (var block in model.Blocks) {
            token.ThrowIfCancellationRequested(); var parts = DocumentReaderEngine.SplitAdapterProjection(block.Text, reader.MaxChars);
            for (int index = 0; index < parts.Count; index++) {
                token.ThrowIfCancellationRequested(); var location = Map(block.Location);
                string id = block.Id + "-part-" + index.ToString(CultureInfo.InvariantCulture); location.BlockAnchor = id;
                var chunk = new ReaderChunk { Id = id, Kind = ReaderInputKind.Xps, Location = location,
                    Text = parts[index], Markdown = parts[index], ContinuesPreviousChunk = index > 0,
                    SourceId = source.SourceId, SourceHash = source.SourceHash,
                    SourceLastWriteUtc = source.LastWriteUtc, SourceLengthBytes = source.LengthBytes,
                    TokenEstimate = Math.Max(1, (parts[index].Length + 3) / 4) };
                if (reader.ComputeHashes) chunk.ChunkHash = DocumentReaderEngine.ComputeChunkHash(chunk);
                chunks.Add(chunk);
            }
        }
        var blocks = model.Blocks.ToDictionary(b => b, b => new OfficeDocumentBlock {
            Id = b.Id, Kind = b.Kind, Text = b.Text, Location = Map(b.Location)
        });
        var tables = model.Tables.ToDictionary(t => t, t => {
            var rows = t.Rows.Take(reader.MaxTableRows).ToArray();
            return new ReaderTable { Kind = t.Kind, Columns = t.Columns, Rows = rows, TotalRowCount = t.TotalRowCount,
                Truncated = rows.Length < t.TotalRowCount, Location = t.Location == null ? null : Map(t.Location),
                ColumnProfiles = ReaderTableProfiler.CreateProfiles(t.Columns, rows) };
        });
        var assets = model.Assets.ToDictionary(a => a, a => new OfficeDocumentAsset {
            Id = a.Id, Kind = a.Kind, MediaType = a.MediaType, Extension = a.Extension, SourceObjectId = a.SourceObjectId,
            PayloadBytes = a.PayloadBytes, PayloadHash = a.PayloadHash, LengthBytes = a.LengthBytes, Location = Map(a.Location)
        });
        var links = model.Links.ToDictionary(link => link, link => new OfficeDocumentLink { Id = link.Id, Kind = link.Kind,
            Uri = link.Uri, DestinationName = link.DestinationName, DestinationPageNumber = link.DestinationPageNumber,
            Text = link.Text, Location = Map(link.Location) });
        var represented = PageTableCoverage(model, tables, token);
        var pages = model.Pages.Select(p => new OfficeDocumentPage { Number = p.Number, Name = p.Name,
            Width = p.Width, Height = p.Height, Location = Map(p.Location),
            Blocks = p.Blocks.Where(b => !represented.Blocks.Contains(b.Id)).Select(b => blocks[b]).ToArray(),
            Tables = p.Tables.Where(t => !represented.Tables.Contains(t.Location!.TableIndex!.Value)).Select(t => tables[t]).ToArray(), Assets = p.Assets.Select(a => assets[a]).ToArray(),
            Links = p.Links.Select(link => links[link]).ToArray() }).ToArray();
        var result = DocumentReaderEngine.CreateRichDocumentResult(chunks, ReaderInputKind.Xps, source,
            model.CapabilitiesUsed.Concat(new[] { "officeimo.reader.xps.pages.native" }), blocks.Values.ToArray(), tables.Values.ToArray(),
            pages, assets.Values.ToArray(), links.Values.ToArray());
        result.Diagnostics = model.Diagnostics.Select(d => new OfficeDocumentDiagnostic { Code = d.Code, Message = d.Message,
            Source = d.Source, Category = OfficeDocumentDiagnosticCategory.Content, Severity = OfficeDocumentDiagnosticSeverity.Warning,
            IsRecoverable = d.IsRecoverable, Location = d.Location == null ? null : Map(d.Location) }).ToArray();
        if (chunks.Count > 0) chunks[0].Warnings = result.Diagnostics.Select(d => d.Message).ToArray();
        return result;
    }

    private static ReaderLocation Map(OfficeDocumentModelLocation location) => new() { Path = location.Path,
        LogicalOrder = location.LogicalOrder,
        Page = location.Page, SourceBlockIndex = location.SourceBlockIndex, SourceBlockKind = location.SourceBlockKind,
        BlockIndex = location.BlockIndex, BlockAnchor = location.BlockAnchor, TableIndex = location.TableIndex };
}
