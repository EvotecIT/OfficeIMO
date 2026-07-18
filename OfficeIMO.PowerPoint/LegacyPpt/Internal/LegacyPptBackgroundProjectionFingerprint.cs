using System.Security.Cryptography;
using System.Text;
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
            var fingerprint = new StringBuilder(background.OuterXml);
            foreach (A.Blip blip in background.Descendants<A.Blip>()) {
                string? relationshipId = blip.Embed?.Value;
                fingerprint.Append("\nimage:").Append(relationshipId);
                if (string.IsNullOrWhiteSpace(relationshipId)) continue;
                try {
                    if (ownerPart.GetPartById(relationshipId!)
                            is not ImagePart imagePart) {
                        fingerprint.Append("|not-image");
                        continue;
                    }
                    fingerprint.Append('|').Append(imagePart.ContentType)
                        .Append('|');
                    using Stream stream = imagePart.GetStream(FileMode.Open,
                        FileAccess.Read);
                    using SHA256 sha256 = SHA256.Create();
                    fingerprint.Append(Convert.ToBase64String(
                        sha256.ComputeHash(stream)));
                } catch (Exception exception) when (exception
                    is ArgumentOutOfRangeException
                    or InvalidDataException
                    or IOException
                    or NotSupportedException
                    or UnauthorizedAccessException) {
                    fingerprint.Append("|unreadable:")
                        .Append(exception.GetType().Name);
                }
            }
            return fingerprint.ToString();
        }
    }
}
