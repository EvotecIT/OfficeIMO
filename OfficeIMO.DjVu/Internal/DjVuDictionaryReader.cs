namespace OfficeIMO.DjVu;

// A dictionary inherits its own component's dictionary, rather than whichever
// Djbz happened to be visited previously while flattening page annotations.
internal static class DjVuDictionaryReader {
    internal static IReadOnlyList<DjVuChunk> Chain(DjVuDocument document, DjVuComponent page, CancellationToken token) {
        var cache = new Dictionary<string, IReadOnlyList<DjVuChunk>>(StringComparer.Ordinal);
        return Resolve(page, 1);
        IReadOnlyList<DjVuChunk> Resolve(DjVuComponent component, int depth) {
            token.ThrowIfCancellationRequested();
            if (depth > document.ReadOptions.MaxDepth) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxDepth));
            if (cache.TryGetValue(component.Id, out var saved)) return saved;
            IReadOnlyList<DjVuChunk> inherited = Array.Empty<DjVuChunk>();
            foreach (var include in component.Form.Children.Where(c => c.Id == "INCL")) {
                string id = DjVuBinary.Utf8.GetString(include.Source, include.Offset, include.Length);
                var candidate = Resolve(document.Components[id], depth + 1);
                if (candidate.Count == 0) continue;
                if (inherited.Count != 0 && !ReferenceEquals(inherited[inherited.Count - 1], candidate[candidate.Count - 1]))
                    throw new InvalidDataException("DjVu component includes conflicting shared dictionaries.");
                inherited = candidate;
            }
            var own = component.Form.Children.Where(c => c.Id == "Djbz").ToArray();
            if (own.Length > 1) throw new InvalidDataException("Multiple dictionaries in a DjVu component.");
            var chain = own.Length == 0 ? inherited : inherited.Concat(own).ToArray();
            cache.Add(component.Id, chain);
            return chain;
        }
    }
}
