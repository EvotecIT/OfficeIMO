namespace OfficeIMO.Reader.Rtf;

internal static partial class RtfReaderAdapter {
    private static RtfReadOptions GetEffectiveReadOptions(ReaderOptions options, ReaderRtfOptions formatOptions) {
        RtfReadOptions result = ReaderRtfOptions.CloneReadOptions(formatOptions.RtfReadOptions)!;
        if (options.Projection == ReaderProjection.TextAndMetadata) result.DeferBinaryPayloads = true;
        return result;
    }

    private static IEnumerable<ReaderChunk> ReadTextAndMetadataChunks(RtfDocument document, SourceMetadata source,
        ReaderOptions options, ReaderRtfOptions formatOptions, IReadOnlyList<OfficeIMO.Rtf.Diagnostics.RtfDiagnostic> diagnostics,
        CancellationToken token) {
        var blocks = new List<(string Kind, int Index, string Text)>();
        for (int index = 0; index < document.Blocks.Count; index++) {
            token.ThrowIfCancellationRequested();
            IRtfBlock block = document.Blocks[index];
            string text = GetNativeBlockText(block, options);
            if (!string.IsNullOrWhiteSpace(text)) blocks.Add((ResolveRtfBlockKind(block), index, text));
        }
        int storyIndex = document.Blocks.Count;
        if (formatOptions.IncludeHeadersAndFooters) foreach (RtfHeaderFooter story in document.HeaderFooters) {
            token.ThrowIfCancellationRequested(); blocks.Add(("header-footer", storyIndex++, story.ToPlainText()));
        }
        if (formatOptions.IncludeNotes) foreach (RtfNote note in document.Notes) {
            token.ThrowIfCancellationRequested(); blocks.Add(("note", storyIndex++, note.ToPlainText()));
        }
        if (!formatOptions.ChunkByBlock && blocks.Count > 0)
            blocks = new List<(string, int, string)> { ("document", 0, string.Join(Environment.NewLine, blocks.Select(block => block.Text))) };
        int emitted = 0;
        foreach (var block in blocks) {
            token.ThrowIfCancellationRequested();
            IReadOnlyList<string> parts = DocumentReaderEngine.SplitAdapterProjection(block.Text, options.MaxChars);
            for (int part = 0; part < parts.Count; part++) yield return EnrichChunk(new ReaderChunk {
                Id = BuildChunkId(block.Kind, block.Index, part, parts.Count), Kind = ReaderInputKind.Rtf,
                Text = parts[part], ContinuesPreviousChunk = part > 0,
                Location = new ReaderLocation { Path = source.Path, BlockIndex = emitted++, SourceBlockIndex = block.Index,
                    SourceBlockKind = block.Kind, BlockAnchor = BuildBlockAnchor(block.Kind, block.Index, part, parts.Count) },
                Warnings = part == 0 ? BuildDiagnosticWarnings(diagnostics) : null
            }, source, options.ComputeHashes);
        }
    }

    private static OfficeDocumentReadResult BuildTextAndMetadataResult(RtfDocument document,
        IReadOnlyList<OfficeIMO.Rtf.Diagnostics.RtfDiagnostic> diagnostics, SourceMetadata source, ReaderOptions options,
        ReaderRtfOptions formatOptions, CancellationToken token) {
        var included = formatOptions.IncludeDiagnostics ? diagnostics : Array.Empty<OfficeIMO.Rtf.Diagnostics.RtfDiagnostic>();
        ReaderChunk[] chunks = Read(document, source, options, formatOptions, included, token).ToArray();
        var result = DocumentReaderEngine.CreateDocumentResult(chunks, ReaderInputKind.Rtf,
            new OfficeDocumentSource { Path = source.Path, SourceId = source.SourceId, SourceHash = source.SourceHash,
                LastWriteUtc = source.LastWriteUtc, LengthBytes = source.LengthBytes, Title = document.Info.Title,
                Author = document.Info.Author, Subject = document.Info.Subject, Keywords = document.Info.Keywords },
            new[] { "officeimo.reader.rtf.text-and-metadata", "officeimo.rtf.semantic-model" });
        result.Metadata = new[] {
            RtfMetadata("rtf-font-count", "FontCount", document.Fonts.Count),
            RtfMetadata("rtf-style-count", "StyleCount", document.Styles.Count),
            RtfMetadata("rtf-source-block-count", "SourceBlockCount", document.Blocks.Count),
            new OfficeDocumentMetadataEntry { Id = "rtf-binary-projection", Category = "rtf.projection", Name = "BinaryPayloads",
                Value = "not-requested", ValueType = "coverage", Attributes = ReaderInspectionMetadataBuilder.Attributes(
                    ("nativeProfile", "source-backed deferred parsing; inert original-source payload sessions are available on RtfReadResult")) }
        };
        result.Diagnostics = result.Diagnostics.Concat(MapRtfDiagnostics(included, source.Path))
            .Concat(MapReaderRtfConversionDiagnostics(formatOptions.Report, source.Path)).ToArray();
        if (formatOptions.IncludePageLocations) {
            var projection = new RtfRichProjection();
            // Page membership is native structure, independent of MaxChars/ChunkByBlock output.
            // Include empty paragraphs so explicit breaks without surrounding text survive too.
            for (int index = 0; index < document.Blocks.Count; index++) {
                token.ThrowIfCancellationRequested();
                IRtfBlock block = document.Blocks[index];
                string kind = ResolveRtfBlockKind(block);
                string id = "rtf-" + kind + "-" + index.ToString("D4", CultureInfo.InvariantCulture);
                projection.Blocks.Add(new OfficeDocumentBlock {
                    Id = id, Kind = kind, Text = GetNativeBlockText(block, options),
                    Location = BuildRtfRichLocation(source.Path, index, kind, id)
                });
            }
            result.Pages = BuildRtfPages(document, projection, source.Path, token);
            result.Diagnostics = result.Diagnostics.Concat(BuildRtfPageDiagnostics(document, result.Pages, source.Path)).ToArray();
        }
        if (options.Inspection == null) return result;
        var inspection = new ReaderInspectionMetadataBuilder(options, "rtf", token);
        if (inspection.Options.IncludeContentSafety)
            OfficeIMO.Reader.FormatInternals.NativeInspectionProjection.InspectContentSafety(inspection,
                safety => document.InspectLoadedContentSafety(safety, token), source.Path);
        inspection.Coverage("rich-content", "not-requested", "Text and native metadata were requested; secondary Markdown, image/object assets, visual and form projections were omitted.", new ReaderLocation { Path = source.Path });
        return inspection.Complete(result);
    }

    private static string GetNativeBlockText(IRtfBlock block, ReaderOptions options) => block switch {
        RtfParagraph paragraph => paragraph.ToPlainText(),
        RtfTable table => string.Join(Environment.NewLine, table.Rows.Take(options.MaxTableRows > 0 ? options.MaxTableRows : 200)
            .Select(row => string.Join("\t", row.Cells.Select(CellText)))),
        RtfObject embedded => embedded.ToPlainText(), RtfShape shape => shape.ToPlainText(), _ => string.Empty
    };
}
