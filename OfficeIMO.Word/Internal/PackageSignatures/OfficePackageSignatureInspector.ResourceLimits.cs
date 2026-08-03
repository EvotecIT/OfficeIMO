#nullable enable
using OfficeIMO.Security;
using System.Xml.Linq;

namespace OfficeIMO.Word {
    internal sealed class OfficePackageDigestWorkBudget {
        private long _remainingBytes;

        internal OfficePackageDigestWorkBudget(long maxBytes) {
            MaxBytes = maxBytes;
            _remainingBytes = maxBytes;
        }

        internal long MaxBytes { get; }

        internal void Reserve(long bytes) {
            if (bytes > _remainingBytes) {
                throw new InvalidDataException("Authenticated package references exceed the " + MaxBytes + " byte aggregate digest-work limit.");
            }
            _remainingBytes -= bytes;
        }
    }

    internal static partial class OfficePackageSignatureInspector {
        private static void ValidateDigestWorkBudget(
            IReadOnlyList<XElement> references,
            int authenticatedReferenceCount,
            OfficePackageSignatureArchive? archive,
            int maxSignedReferences,
            OfficePackageDigestWorkBudget digestWorkBudget) {
            if (authenticatedReferenceCount > maxSignedReferences) {
                throw new InvalidDataException("The XML signature contains more than " + maxSignedReferences + " authenticated references.");
            }
            if (archive == null) return;

            long signatureDigestBytes = 0;
            XNamespace ds = "http://www.w3.org/2000/09/xmldsig#";
            foreach (XElement reference in references) {
                string? targetPartUri = NormalizePackagePartReference((string?)reference.Attribute("URI"));
                if (targetPartUri == null || !archive.TryGetPartLength(targetPartUri, out long partLength)) continue;
                long transformCount = reference
                    .Element(ds + "Transforms")?
                    .Elements(ds + "Transform")
                    .LongCount() ?? 0;
                long workPasses = checked(transformCount + 1);
                if (partLength > 0 && workPasses > long.MaxValue / partLength) {
                    throw new InvalidDataException("Authenticated package references exceed the " + digestWorkBudget.MaxBytes + " byte aggregate digest-work limit.");
                }
                long referenceDigestBytes = partLength * workPasses;
                if (referenceDigestBytes > long.MaxValue - signatureDigestBytes) {
                    throw new InvalidDataException("Authenticated package references exceed the " + digestWorkBudget.MaxBytes + " byte aggregate digest-work limit.");
                }
                signatureDigestBytes += referenceDigestBytes;
            }
            digestWorkBudget.Reserve(signatureDigestBytes);
        }
    }
}
