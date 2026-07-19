using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    public static IEnumerable<ReaderChunk> Read(
        Stream stream,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        ValidateReadableStream(stream);
        ReaderOptions effective = NormalizeOptions(options);
        string logicalName = NormalizeLogicalSourceName(sourceName, "memory");
        Stream readStream = ReaderInputLimits.EnsureSeekableReadStream(
            stream,
            ResolveStreamMaxInputBytes(logicalName, effective, stream.CanSeek),
            cancellationToken,
            out bool ownsReadStream);
        try {
            if (!TryResolveStreamHandler(
                    readStream,
                    logicalName,
                    effective,
                    cancellationToken,
                    out ReaderHandlerDescriptor handler,
                    out ReaderDetectionResult detection)) {
                throw CreateUnsupportedInputException(logicalName, detection);
            }

            long position = readStream.Position;
            SourceInfo source = BuildSourceInfoFromStream(readStream, logicalName,
                ShouldComputeSourceHash(handler, effective), cancellationToken);
            readStream.Position = position;
            IEnumerable<ReaderChunk> chunks;
            if (handler.ReadStream != null) {
                chunks = handler.ReadStream(readStream, logicalName, effective, cancellationToken)
                    ?? throw new InvalidOperationException($"Reader handler '{handler.Id}' returned null chunks.");
            } else if (handler.ReadDocumentStream != null) {
                OfficeDocumentReadResult result = ValidateDocumentResult(
                    handler.ReadDocumentStream(readStream, logicalName, effective, cancellationToken),
                    handler.Id);
                chunks = result.Chunks ?? Array.Empty<ReaderChunk>();
            } else if (handler.ReadDocumentStreamAsync != null) {
                throw CreateAsyncOnlyHandlerException(handler.Id, "stream");
            } else {
                throw new NotSupportedException($"Reader handler '{handler.Id}' does not support stream input.");
            }

            return chunks
                .Select(chunk => EnrichChunk(chunk, source, effective.ComputeHashes))
                .ToArray();
        } finally {
            if (ownsReadStream) readStream.Dispose();
        }
    }

    public static IEnumerable<ReaderChunk> Read(
        byte[] bytes,
        string? sourceName = null,
        ReaderOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return Read(stream, sourceName, options, cancellationToken).ToArray();
    }

    private static NotSupportedException CreateUnsupportedInputException(string sourceName, ReaderDetectionResult detection) {
        string kind = detection.Kind == ReaderInputKind.Unknown ? "unknown" : detection.Kind.ToString();
        return new NotSupportedException(
            $"No registered Reader handler accepts '{sourceName}' (detected kind: {kind}). " +
            "Install the matching OfficeIMO.Reader.<Format> package and add its handler to OfficeDocumentReaderBuilder.");
    }
}
