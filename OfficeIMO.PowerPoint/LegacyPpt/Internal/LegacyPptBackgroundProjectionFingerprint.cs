using System.Security.Cryptography;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Fingerprints background XML together with directly related image bytes so
    /// in-place image part replacements cannot be mistaken for no-op edits.
    /// </summary>
    internal static class LegacyPptBackgroundProjectionFingerprint {
        /// <summary>The fingerprint of the background of a slide-like part (its <c>p:cSld/p:bg</c>).</summary>
        internal static string Create(LegacyPptPartView ownerPart) =>
            Create(ownerPart, ownerPart.Root.Child(LegacyPptXml.P, "cSld").Child(LegacyPptXml.P, "bg"));

        internal static string Create(LegacyPptPartView ownerPart,
            OfficeXmlNode background) {
            if (ownerPart.IsNone) throw new ArgumentNullException(nameof(ownerPart));
            if (background.IsNone) return string.Empty;
            using IncrementalHash fingerprint = LegacyPptProjectionDigest
                .CreateBuilder();
            LegacyPptProjectionDigest.Append(fingerprint,
                LegacyPptXml.Digest(background));
            foreach (OfficeXmlNode blip in background.Descendants(LegacyPptXml.A, "blip")) {
                string? relationshipId = blip.Attribute(LegacyPptXml.R, "embed");
                LegacyPptProjectionDigest.Append(fingerprint,
                    "image:" + relationshipId);
                if (string.IsNullOrWhiteSpace(relationshipId)) continue;
                try {
                    if (!ownerPart.TryGetRelationship(relationshipId,
                            out LegacyPptRelationshipView relationship)
                        || relationship.IsExternal || relationship.TargetPart.IsNone) {
                        LegacyPptProjectionDigest.Append(fingerprint,
                            "unreadable:" + nameof(ArgumentOutOfRangeException));
                        continue;
                    }
                    if (!relationship.IsType(LegacyPptPackageView.ImageRelationshipType)) {
                        LegacyPptProjectionDigest.Append(fingerprint,
                            "not-image");
                        continue;
                    }
                    LegacyPptPartView imagePart = relationship.TargetPart;
                    LegacyPptProjectionDigest.Append(fingerprint,
                        imagePart.ContentType);
                    LegacyPptProjectionDigest.Append(fingerprint,
                        HashBytes(imagePart.GetBytes()));
                } catch (Exception exception) when (exception
                    is InvalidDataException
                    or IOException
                    or NotSupportedException
                    or UnauthorizedAccessException) {
                    LegacyPptProjectionDigest.Append(fingerprint,
                        "unreadable:" + exception.GetType().Name);
                }
            }
            return LegacyPptProjectionDigest.Finish(fingerprint);
        }

        /// <summary>The Base64 SHA-256 of a part's bytes.</summary>
        internal static string HashBytes(byte[] bytes) {
#if NET5_0_OR_GREATER
            return Convert.ToBase64String(SHA256.HashData(bytes));
#else
            using SHA256 sha256 = SHA256.Create();
            return Convert.ToBase64String(sha256.ComputeHash(bytes));
#endif
        }
    }
}
