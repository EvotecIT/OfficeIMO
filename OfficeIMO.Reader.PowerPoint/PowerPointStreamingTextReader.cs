using OfficeIMO.PowerPoint;

namespace OfficeIMO.Reader.PowerPoint;

/// <summary>
/// Reads the text projection of a PPTX package with <see cref="PowerPointSlideReader"/>, one slide part at a time,
/// instead of loading the presentation model. It maps each slide to the blocks, tables and chunks the model mapping
/// produces and lets the read scope reduce them to text, so the text is the same. Packages it cannot read that way
/// (binary or encrypted input, input over the configured limits, malformed parts, tables without a column grid)
/// make it decline; the caller then reads the model, which reports any error as before.
/// </summary>
internal static class PowerPointStreamingTextReader {
    internal const string Capability = "officeimo.reader.powerpoint.streamed-text";

    /// <summary>Whether the read asked for the text projection and nothing this pass cannot honor.</summary>
    internal static bool Applies(ReaderOptions readerOptions) =>
        DocumentReaderEngine.IsTextProjection
        && string.IsNullOrEmpty(readerOptions.OpenPassword)
        && readerOptions.Inspection == null;

    /// <summary>Returns the streamed result, or null to read the presentation through the model.</summary>
    internal static OfficeDocumentReadResult? TryRead(Stream stream, string sourceName, ReaderOptions readerOptions,
        ReaderPowerPointOptions options, CancellationToken cancellationToken) {
        if (!stream.CanSeek) return null;
        long start = stream.Position;
        try {
            PowerPointLoadOptions loadOptions = PowerPointReaderAdapter.CreateLoadOptions(readerOptions, sourceName);
            loadOptions.Password = null;
            using PowerPointSlideReader reader = PowerPointSlideReader.Open(stream, loadOptions, cancellationToken);
            PowerPointExtractionExtensions.PowerPointExtractOptions extract = PowerPointExtractionExtensions.Snapshot(
                new PowerPointExtractionExtensions.PowerPointExtractOptions {
                    IncludeNotes = options.IncludeNotes,
                    IncludeTables = options.IncludeTables,
                    IncludeHiddenShapes = options.IncludeHiddenShapes,
                    IncludeMarkdown = false
                });
            int maxChars = PowerPointExtractionExtensions.ResolveMaxChars(new PowerPointExtractChunkingOptions { MaxChars = readerOptions.MaxChars });
            var chunks = new List<ReaderChunk>();
            var projection = new StreamedSlideProjection();
            foreach (PowerPointSlideContent slide in reader.ReadSlides(options.IncludeNotes, cancellationToken)) {
                foreach (PowerPointExtractChunk chunk in PowerPointExtractionExtensions.ExtractSlideChunks(slide, extract, maxChars, sourceName, cancellationToken)) {
                    chunks.Add(PowerPointReaderAdapter.ToReaderChunk(chunk, chunks.Count, sourceName, null));
                }
                if (!PowerPointRichMapping.AddStreamedSlide(slide, sourceName, readerOptions, options, projection)) return null;
            }

            string[] capabilities = { OfficeDocumentReaderBuilderPowerPointExtensions.HandlerId, Capability };
            OfficeDocumentReadResult result = DocumentReaderEngine.CreateStreamedTextResult(chunks, ReaderInputKind.PowerPoint, capabilities);
            return PowerPointRichMapping.CompleteStreamed(result, reader.ReadProperties(), projection, capabilities);
        } catch (OperationCanceledException) {
            throw;
        } catch (Exception exception) when (exception is not OutOfMemoryException) {
            // The model load reports damaged, unusual or over-limit packages with its own errors.
            return null;
        } finally {
            stream.Position = start;
        }
    }
}
