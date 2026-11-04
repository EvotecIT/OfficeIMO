using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {

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
