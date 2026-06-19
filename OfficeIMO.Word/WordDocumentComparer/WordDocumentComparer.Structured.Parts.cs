using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public static partial class WordDocumentComparer {
        private static Dictionary<HeaderPart, string> CreateHeaderPartKeys(MainDocumentPart mainPart) {
            var keys = new Dictionary<HeaderPart, string>();
            var typeOrdinals = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (HeaderReference reference in mainPart.Document?.Descendants<HeaderReference>() ?? Enumerable.Empty<HeaderReference>()) {
                if (reference.Id?.Value is not string relationshipId) {
                    continue;
                }

                if (mainPart.GetPartById(relationshipId) is not HeaderPart headerPart || keys.ContainsKey(headerPart)) {
                    continue;
                }

                string typeKey = GetHeaderFooterReferenceTypeKey(reference.Type?.Value);
                int ordinal = GetAndIncrementOrdinal(typeOrdinals, typeKey);
                keys[headerPart] = HeaderPartKeyPrefix + typeKey + ":" + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            return keys;
        }

        private static Dictionary<FooterPart, string> CreateFooterPartKeys(MainDocumentPart mainPart) {
            var keys = new Dictionary<FooterPart, string>();
            var typeOrdinals = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (FooterReference reference in mainPart.Document?.Descendants<FooterReference>() ?? Enumerable.Empty<FooterReference>()) {
                if (reference.Id?.Value is not string relationshipId) {
                    continue;
                }

                if (mainPart.GetPartById(relationshipId) is not FooterPart footerPart || keys.ContainsKey(footerPart)) {
                    continue;
                }

                string typeKey = GetHeaderFooterReferenceTypeKey(reference.Type?.Value);
                int ordinal = GetAndIncrementOrdinal(typeOrdinals, typeKey);
                keys[footerPart] = FooterPartKeyPrefix + typeKey + ":" + ordinal.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            return keys;
        }

        private static string GetHeaderPartKey(IReadOnlyDictionary<HeaderPart, string> keys, HeaderPart headerPart, int fallbackIndex) {
            return keys.TryGetValue(headerPart, out string? key)
                ? key
                : HeaderPartKeyPrefix + "unreferenced:" + fallbackIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string GetFooterPartKey(IReadOnlyDictionary<FooterPart, string> keys, FooterPart footerPart, int fallbackIndex) {
            return keys.TryGetValue(footerPart, out string? key)
                ? key
                : FooterPartKeyPrefix + "unreferenced:" + fallbackIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string GetHeaderFooterReferenceTypeKey(HeaderFooterValues? type) {
            if (type == HeaderFooterValues.First) {
                return "first";
            }

            if (type == HeaderFooterValues.Even) {
                return "even";
            }

            return "default";
        }

        private static int GetAndIncrementOrdinal(Dictionary<string, int> ordinals, string typeKey) {
            ordinals.TryGetValue(typeKey, out int ordinal);
            ordinals[typeKey] = ordinal + 1;
            return ordinal;
        }

        private static double GetImageSimilarity(ImageSnapshot source, ImageSnapshot target) {
            if (!string.Equals(source.PartKey, target.PartKey, StringComparison.Ordinal)) {
                return 0;
            }

            bool sameVisualSignature = string.Equals(source.VisualSignature, target.VisualSignature, StringComparison.Ordinal);
            if (source.ExternalUri != null || target.ExternalUri != null) {
                if (string.Equals(source.ExternalUri, target.ExternalUri, StringComparison.Ordinal)) {
                    return sameVisualSignature ? 1 : 0.8;
                }

                return sameVisualSignature ? 0.6 : 0;
            }

            if (source.EmbeddedFingerprint != null &&
                target.EmbeddedFingerprint != null &&
                source.EmbeddedFingerprint.Equals(target.EmbeddedFingerprint)) {
                return sameVisualSignature ? 1 : 0.8;
            }

            return sameVisualSignature ? 0.6 : 0;
        }
    }
}
