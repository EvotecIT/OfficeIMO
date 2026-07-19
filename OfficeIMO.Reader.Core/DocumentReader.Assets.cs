using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    /// <summary>
    /// Reads a supported document file and returns discovered assets in source order.
    /// </summary>
    /// <param name="path">Source file path.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<OfficeDocumentAsset> ReadAssets(string path, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        return ExtractAssets(ReadDocument(path, options, cancellationToken), cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Reads a supported document stream and returns discovered assets in source order.
    /// </summary>
    /// <param name="stream">Source stream. This method does not close the stream.</param>
    /// <param name="sourceName">Optional source name used for kind detection and citations.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<OfficeDocumentAsset> ReadAssets(Stream stream, string? sourceName = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        return ExtractAssets(ReadDocument(stream, sourceName, options, cancellationToken), cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Reads a supported document from bytes and returns discovered assets in source order.
    /// </summary>
    /// <param name="bytes">Source bytes.</param>
    /// <param name="sourceName">Optional source name used for kind detection and citations.</param>
    /// <param name="options">Extraction options.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<OfficeDocumentAsset> ReadAssets(byte[] bytes, string? sourceName = null, ReaderOptions? options = null, CancellationToken cancellationToken = default) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        using var stream = new MemoryStream(bytes, writable: false);
        return ReadAssets(stream, sourceName, options, cancellationToken);
    }

    /// <summary>
    /// Returns assets already attached to a shared read result, optionally filtered by caller-owned policy.
    /// </summary>
    /// <param name="result">Read result to inspect.</param>
    /// <param name="predicate">Optional predicate used to select assets.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    public static IReadOnlyList<OfficeDocumentAsset> ExtractAssets(OfficeDocumentReadResult result, Func<OfficeDocumentAsset, bool>? predicate = null, CancellationToken cancellationToken = default) {
        if (result == null) throw new ArgumentNullException(nameof(result));

        IReadOnlyList<OfficeDocumentAsset> assets = result.Assets ?? Array.Empty<OfficeDocumentAsset>();
        if (assets.Count == 0) {
            return Array.Empty<OfficeDocumentAsset>();
        }

        if (predicate == null) {
            cancellationToken.ThrowIfCancellationRequested();
            return assets;
        }

        List<OfficeDocumentAsset>? selected = null;
        for (int i = 0; i < assets.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeDocumentAsset asset = assets[i];
            if (predicate(asset)) {
                selected ??= new List<OfficeDocumentAsset>();
                selected.Add(asset);
            }
        }

        return selected == null || selected.Count == 0 ? Array.Empty<OfficeDocumentAsset>() : selected.ToArray();
    }
}
