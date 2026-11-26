using OfficeIMO.Epub;

namespace OfficeIMO.Reader.Epub;

internal static partial class EpubReaderAdapter {
    internal static ReaderHandlerRegistration CreateImportHandler(string format, string[] extensions, ReaderInputKind kind,
        long maxInputBytes, Func<Stream, string, long, CancellationToken, OfficeConversionResult<EpubPublication>> import) {
        OfficeDocumentReadResult ReadPath(string path, ReaderOptions options, CancellationToken token) {
            using var input = File.OpenRead(path);
            return ReadImport(input, BuildSourceMetadataFromPath(path, options.ComputeHashes), options, token);
        }

        OfficeDocumentReadResult ReadStream(Stream stream, string? name, ReaderOptions options, CancellationToken token) {
            string sourceName = string.IsNullOrWhiteSpace(name) ? "document" + extensions[0] : name!;
            Stream input = ReaderInputLimits.EnsureSeekableReadStream(stream, Math.Min(options.MaxInputBytes ?? maxInputBytes, maxInputBytes),
                token, out bool ownsInput);
            try {
                SourceMetadata source = BuildSourceMetadataFromStream(input, sourceName, options.ComputeHashes);
                input.Position = 0;
                return ReadImport(input, source, options, token);
            } finally {
                if (ownsInput) input.Dispose();
            }
        }

        OfficeDocumentReadResult ReadImport(Stream input, SourceMetadata source, ReaderOptions options, CancellationToken token) {
            var converted = import(input, source.Path, Math.Min(options.MaxInputBytes ?? maxInputBytes, maxInputBytes), token);
            EpubDocument book = EpubImportedPublication.Open(converted.Value, token);
            OfficeDocumentReadResult result = BuildEpubDocumentResult(book, source, options, token, kind);
            result.CapabilitiesUsed = result.CapabilitiesUsed.Concat(new[] { "officeimo.reader." + format }).ToArray();
            result.Diagnostics = result.Diagnostics.Concat(converted.Report.Diagnostics.Select(item => new OfficeDocumentDiagnostic {
                Code = item.Code,
                Message = item.Message,
                Source = item.Format,
                Severity = item.Severity switch {
                    OfficeDiagnosticSeverity.Information => OfficeDocumentDiagnosticSeverity.Information,
                    OfficeDiagnosticSeverity.Error => OfficeDocumentDiagnosticSeverity.Error,
                    _ => OfficeDocumentDiagnosticSeverity.Warning
                },
                Attributes = new Dictionary<string, string> { ["lossKind"] = item.LossKind.ToString() }
            })).ToArray();
            return result;
        }

        return new ReaderHandlerRegistration {
            Id = "officeimo.reader." + format,
            Origin = ReaderHandlerOrigin.OfficeIMO,
            Kind = kind,
            Extensions = extensions,
            UseDetectedKindFallback = false,
            DefaultMaxInputBytes = maxInputBytes,
            MaxInputBytesCeiling = maxInputBytes,
            ReadDocumentPath = ReadPath,
            ReadDocumentStream = ReadStream,
            ReadPath = (path, options, token) => ReadPath(path, options, token).Chunks,
            ReadStream = (stream, name, options, token) => ReadStream(stream, name, options, token).Chunks
        };
    }
}
