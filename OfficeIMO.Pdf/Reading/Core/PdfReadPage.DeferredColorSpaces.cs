using System.Collections;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private static readonly IReadOnlyDictionary<string, PdfPageColorSpace> NoColorSpaceResources =
        new Dictionary<string, PdfPageColorSpace>(StringComparer.Ordinal);

    /// <summary>
    /// Returns the visibly invoked color-space resources for text interpretation. Selecting them needs
    /// a visual pass over the content, so that pass runs on the first lookup of a declared resource
    /// name instead of before every read. Content that names no declared color space never pays for it.
    /// </summary>
    private IReadOnlyDictionary<string, PdfPageColorSpace> GetDeferredColorSpaceResources(
        PdfDictionary? resources,
        Func<PdfPageInvokedResourceNames> invokedResourceNames,
        PageContentBudget pageContentBudget) {
        if (resources == null ||
            !resources.Items.TryGetValue("ColorSpace", out PdfObject? colorSpacesObject) ||
            ResolveDictionary(colorSpacesObject) is not PdfDictionary colorSpaces ||
            colorSpaces.Items.Count == 0) {
            return NoColorSpaceResources;
        }
        if (colorSpaces.Items.Count > _limits.MaxColorSpaceResourcesPerPage) {
            throw new InvalidDataException(
                $"The page declares more than {_limits.MaxColorSpaceResourcesPerPage} color-space resources.");
        }

        return new DeferredColorSpaceResources(colorSpaces, () =>
            GetColorSpaceResources(resources, invokedResourceNames().ColorSpaces, pageContentBudget));
    }

    private sealed class DeferredColorSpaceResources : IReadOnlyDictionary<string, PdfPageColorSpace> {
        private readonly PdfDictionary _declared;
        private Func<Dictionary<string, PdfPageColorSpace>>? _resolve;
        private Dictionary<string, PdfPageColorSpace>? _resolved;

        internal DeferredColorSpaceResources(PdfDictionary declared, Func<Dictionary<string, PdfPageColorSpace>> resolve) {
            _declared = declared;
            _resolve = resolve;
        }

        private Dictionary<string, PdfPageColorSpace> Resolved {
            get {
                if (_resolved == null) {
                    _resolved = _resolve!();
                    _resolve = null;
                }
                return _resolved;
            }
        }

        public bool TryGetValue(string key, out PdfPageColorSpace value) {
            if (!_declared.Items.ContainsKey(key)) {
                value = default;
                return false;
            }
            return Resolved.TryGetValue(key, out value);
        }

        public bool ContainsKey(string key) => _declared.Items.ContainsKey(key) && Resolved.ContainsKey(key);
        public PdfPageColorSpace this[string key] => TryGetValue(key, out PdfPageColorSpace value) ? value : throw new KeyNotFoundException(key);
        public IEnumerable<string> Keys => Resolved.Keys;
        public IEnumerable<PdfPageColorSpace> Values => Resolved.Values;
        public int Count => Resolved.Count;
        public IEnumerator<KeyValuePair<string, PdfPageColorSpace>> GetEnumerator() => Resolved.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
