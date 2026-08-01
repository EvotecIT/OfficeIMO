#nullable enable
using System.Xml.Linq;

namespace OfficeIMO.Word {
    internal static partial class OfficePackageSignatureInspector {
        private static void ValidateDigestWorkBudget(
            IReadOnlyList<XElement> references,
            int authenticatedReferenceCount,
            OfficePackageSignatureArchive? archive,
            int maxSignedReferences,
            long maxTotalDigestBytes) {
            if (authenticatedReferenceCount > maxSignedReferences) {
                throw new InvalidDataException("The XML signature contains more than " + maxSignedReferences + " authenticated references.");
            }
            if (archive == null) return;

            long totalDigestBytes = 0;
            foreach (XElement reference in references) {
                string? targetPartUri = NormalizePackagePartReference((string?)reference.Attribute("URI"));
                if (targetPartUri == null || !archive.TryGetPartLength(targetPartUri, out long partLength)) continue;
                if (partLength > maxTotalDigestBytes - totalDigestBytes) {
                    throw new InvalidDataException("Authenticated package references exceed the " + maxTotalDigestBytes + " byte aggregate digest-work limit.");
                }
                totalDigestBytes += partLength;
            }
        }
    }
}
