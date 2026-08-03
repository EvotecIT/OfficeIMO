#nullable enable
using System.Security.Cryptography.Xml;
using System.Text;
using System.Xml;

namespace OfficeIMO.Word {
    internal static partial class OfficePackageSignatureValidator {
        private static readonly string[] LocalReferenceIdAttributeNames = { "Id", "ID", "id" };

        private static void EnsureLocalSignedInfoDigestWorkWithinLimit(
            XmlDocument document,
            SignedXml signedXml,
            int certificateCandidateCount,
            long maxTotalDigestBytes) {
            if (certificateCandidateCount <= 0 || signedXml.SignedInfo == null) return;

            long totalDigestWorkBytes = 0;
            var targetSizes = new Dictionary<string, long>(StringComparer.Ordinal);
            IReadOnlyDictionary<string, XmlElement?> targetsById = IndexLocalReferenceTargets(document);
            foreach (object item in signedXml.SignedInfo.References) {
                if (item is not Reference reference) continue;
                string uri = reference.Uri ?? string.Empty;
                if (uri.Length > 0 && !uri.StartsWith("#", StringComparison.Ordinal)) continue;

                if (!targetSizes.TryGetValue(uri, out long targetBytes)) {
                    XmlElement? target = ResolveLocalReferenceTarget(document, targetsById, uri);
                    targetBytes = GetSerializedElementByteCount(target ?? document.DocumentElement);
                    targetSizes.Add(uri, targetBytes);
                }

                int transformPasses = (reference.TransformChain?.Count ?? 0) + 1;
                long referenceWorkBytes = SaturatingMultiply(targetBytes, transformPasses);
                long candidateWorkBytes = SaturatingMultiply(referenceWorkBytes, certificateCandidateCount);
                if (candidateWorkBytes > maxTotalDigestBytes - totalDigestWorkBytes) {
                    throw new InvalidDataException(
                        "Local SignedInfo references exceed the " + maxTotalDigestBytes +
                        " byte aggregate digest-work limit across certificate candidates.");
                }
                totalDigestWorkBytes += candidateWorkBytes;
            }
        }

        private static XmlElement? ResolveLocalReferenceTarget(
            XmlDocument document,
            IReadOnlyDictionary<string, XmlElement?> targetsById,
            string uri) {
            if (uri.Length == 0) return document.DocumentElement;
            if (uri.Length == 1) return null;
            return targetsById.TryGetValue(uri.Substring(1), out XmlElement? target)
                ? target
                : null;
        }

        private static IReadOnlyDictionary<string, XmlElement?> IndexLocalReferenceTargets(XmlDocument document) {
            var targetsById = new Dictionary<string, XmlElement?>(StringComparer.Ordinal);
            if (document.DocumentElement == null) return targetsById;
            foreach (XmlElement element in EnumerateElements(document.DocumentElement)) {
                foreach (string attributeName in LocalReferenceIdAttributeNames) {
                    if (!element.HasAttribute(attributeName)) continue;
                    string id = element.GetAttribute(attributeName);
                    if (id.Length == 0) continue;
                    if (targetsById.TryGetValue(id, out XmlElement? existing)) {
                        if (!ReferenceEquals(existing, element)) targetsById[id] = null;
                    } else {
                        targetsById.Add(id, element);
                    }
                }
            }
            return targetsById;
        }

        private static IEnumerable<XmlElement> EnumerateElements(XmlElement root) {
            yield return root;
            foreach (XmlElement descendant in root.GetElementsByTagName("*").OfType<XmlElement>()) {
                yield return descendant;
            }
        }

        private static long GetSerializedElementByteCount(XmlElement? element) =>
            element == null ? 0L : Encoding.UTF8.GetByteCount(element.OuterXml);

        private static long SaturatingMultiply(long value, int multiplier) =>
            value > long.MaxValue / multiplier ? long.MaxValue : value * multiplier;
    }
}
