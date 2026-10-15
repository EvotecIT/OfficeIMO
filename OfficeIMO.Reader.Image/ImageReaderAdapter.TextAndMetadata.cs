using OfficeIMO.Drawing;

namespace OfficeIMO.Reader.Image;

internal static partial class ImageReaderAdapter {
    private static OfficeDocumentReadResult BuildTextAndMetadataResult(ReaderAdapterInputSnapshot input, OfficeImageInfo info,
        ReaderOptions options, ReaderImageOptions imageOptions, CancellationToken token) {
        string sourceName = input.Source.Path ?? "image.bin";
        var location = new ReaderLocation { Path = sourceName, SourceBlockIndex = 0, SourceBlockKind = "image-metadata", BlockAnchor = "image-metadata" };
        IReadOnlyList<string> parts = DocumentReaderEngine.SplitAdapterProjection(BuildPlainText(sourceName, info, input.Bytes.LongLength), options.MaxChars);
        var chunks = new List<ReaderChunk>(parts.Count);
        for (int index = 0; index < parts.Count; index++) {
            token.ThrowIfCancellationRequested();
            var chunk = new ReaderChunk {
                Id = "image-metadata-" + index.ToString("D4", CultureInfo.InvariantCulture),
                Kind = ReaderInputKind.Unknown, Text = parts[index], ContinuesPreviousChunk = index > 0,
                Location = new ReaderLocation { Path = sourceName, BlockIndex = index, SourceBlockIndex = 0,
                    SourceBlockKind = location.SourceBlockKind, BlockAnchor = location.BlockAnchor }
            };
            DocumentReaderEngine.ApplyAdapterSource(chunk, input, options.ComputeHashes);
            chunks.Add(chunk);
        }
        var result = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Unknown, Source = new OfficeDocumentSource {
                Path = input.Source.Path, SourceId = input.Source.SourceId, SourceHash = input.Source.SourceHash,
                LastWriteUtc = input.Source.LastWriteUtc, LengthBytes = input.Source.LengthBytes
            },
            Chunks = chunks.AsReadOnly(), Metadata = BuildMetadata(info, input.Bytes.LongLength).ToArray(),
            CapabilitiesUsed = new[] { OfficeDocumentReaderBuilderImageExtensions.HandlerId, "officeimo.drawing.image-identification", "officeimo.reader.image.text-and-metadata" }
        };
        result.Source.Title = Path.GetFileName(sourceName);
        if (options.Inspection != null) ProjectMetadataEvidence(input.Bytes, result, options, imageOptions, location, token);
        return result;
    }
}
