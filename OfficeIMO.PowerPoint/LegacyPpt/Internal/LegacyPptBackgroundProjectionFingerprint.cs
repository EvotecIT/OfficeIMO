using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint.LegacyPpt.Internal {
    /// <summary>
    /// Fingerprints background XML together with directly related image bytes so
    /// in-place ImagePart replacements cannot be mistaken for no-op edits.
    /// </summary>
    internal static class LegacyPptBackgroundProjectionFingerprint {
        internal static string Create(OpenXmlPart ownerPart,
            P.Background? background) {
            if (ownerPart == null) throw new ArgumentNullException(
                nameof(ownerPart));
            if (background == null) return string.Empty;
            using IncrementalHash fingerprint = LegacyPptProjectionDigest
                .CreateBuilder();
            LegacyPptProjectionDigest.Append(fingerprint,
                background.OuterXml);
            foreach (A.Blip blip in background.Descendants<A.Blip>()) {
                string? relationshipId = blip.Embed?.Value;
                LegacyPptProjectionDigest.Append(fingerprint,
                    "image:" + relationshipId);
                if (string.IsNullOrWhiteSpace(relationshipId)) continue;
                try {
                    if (ownerPart.GetPartById(relationshipId!)
                            is not ImagePart imagePart) {
                        LegacyPptProjectionDigest.Append(fingerprint,
                            "not-image");
                        continue;
                    }
                    LegacyPptProjectionDigest.Append(fingerprint,
                        imagePart.ContentType);
                    using Stream stream = imagePart.GetStream(FileMode.Open,
                        FileAccess.Read);
                    using SHA256 sha256 = SHA256.Create();
                    LegacyPptProjectionDigest.Append(fingerprint,
                        Convert.ToBase64String(sha256.ComputeHash(stream)));
                } catch (Exception exception) when (exception
                    is ArgumentOutOfRangeException
                    or InvalidDataException
                    or IOException
                    or NotSupportedException
                    or UnauthorizedAccessException) {
                    LegacyPptProjectionDigest.Append(fingerprint,
                        "unreadable:" + exception.GetType().Name);
                }
            }
            return LegacyPptProjectionDigest.Finish(fingerprint);
        }
    }
}
