using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static bool IsRegisteredDirectoryBundle(string path) =>
        TryResolveCustomHandlerByPath(path, out ReaderHandlerDescriptor handler)
        && handler.ReadDirectoryBundle != null;

    private static OfficeDocumentReadResult ReadDirectoryBundle(string path,
        ReaderOptions? options, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!TryResolveCustomHandlerByPath(path, out ReaderHandlerDescriptor handler)
            || handler.ReadDirectoryBundle == null) {
            throw new IOException($"'{path}' is a directory. Use {nameof(ReadFolder)}(...) to ingest directories.");
        }
        ReaderOptions effective = NormalizeOptions(options);
        effective.MaxInputBytes = ResolveSelectedHandlerMaxInputBytes(handler, path, effective)
            ?? DefaultUnidentifiedStreamMaxInputBytes;
        OfficeDocumentReadResult result = ValidateDocumentResult(
            handler.ReadDirectoryBundle(path, effective, cancellationToken), handler.Id);
        cancellationToken.ThrowIfCancellationRequested();
        long? length = result.Source?.LengthBytes;
        if (!length.HasValue || length.Value < 0 || length.Value > effective.MaxInputBytes.Value) {
            throw new InvalidDataException("The directory-package handler must report a nonnegative source byte length within the input limit.");
        }
        SourceInfo source = BuildSourceInfoFromPath(path, computeHash: false, cancellationToken);
        source.LengthBytes = length;
        source.SourceHash = effective.ComputeHashes ? result.Source?.SourceHash : null;
        if (effective.ComputeHashes && string.IsNullOrWhiteSpace(source.SourceHash)) {
            throw new InvalidDataException("The directory-package handler must supply a snapshot hash when ComputeHashes is enabled.");
        }
        return ApplyDetectionDiagnostics(FinalizeHandlerDocumentResult(result, source, effective.ComputeHashes),
            DetectDirectoryBundle(path));
    }

    private static ReaderDetectionResult DetectDirectoryBundle(string path) {
        if (!TryResolveCustomHandlerByPath(path, out ReaderHandlerDescriptor handler)
            || handler.ReadDirectoryBundle == null) {
            throw new IOException($"'{path}' is a directory without a registered document-package handler.");
        }
        ReaderDetectionResult result = BuildExtensionDetection(path);
        result.Evidence = result.Evidence.Concat(new[] { "directory-package-handler:" + handler.Id }).ToArray();
        return result;
    }
}
