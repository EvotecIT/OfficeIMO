using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Reader;

/// <summary>Shared cancellation-aware preflight helpers for deterministic materializers.</summary>
internal static class ReaderMaterializationPreflight {
    internal static List<T> ToList<T>(IEnumerable<T> source, CancellationToken cancellationToken) {
        var results = new List<T>();
        using IEnumerator<T> enumerator = source.GetEnumerator();
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            bool hasNext = enumerator.MoveNext();
            cancellationToken.ThrowIfCancellationRequested();
            if (!hasNext) return results;
            results.Add(enumerator.Current);
        }
    }
}
