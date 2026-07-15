using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    /// <summary>Creates deterministic fingerprints over a presentation package and its relationships.</summary>
    internal static class PowerPointPackageFingerprint {
        internal static string Create(PresentationDocument document,
            Action<OpenXmlPart, OpenXmlElement>? normalizeRoot = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var parts = new HashSet<OpenXmlPart>();
            foreach (IdPartPair pair in document.Parts) CollectParts(pair.OpenXmlPart, parts);

            var content = new StringBuilder();
            foreach (OpenXmlPart part in parts.OrderBy(item => item.Uri.ToString(), StringComparer.Ordinal)) {
                content.Append(part.Uri).Append('|').Append(part.ContentType).Append('|');
                try {
                    OpenXmlPartRootElement? root = part.RootElement;
                    if (root != null) {
                        if (normalizeRoot == null) {
                            content.Append(root.OuterXml);
                        } else {
                            OpenXmlElement normalized = root.CloneNode(true);
                            normalizeRoot(part, normalized);
                            content.Append(normalized.OuterXml);
                        }
                    } else {
                        using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                        using var memory = new MemoryStream();
                        stream.CopyTo(memory);
                        content.Append(Convert.ToBase64String(memory.ToArray()));
                    }
                } catch (InvalidDataException) {
                    content.Append("unreadable");
                }
                foreach (IdPartPair relationship in part.Parts.OrderBy(item => item.RelationshipId,
                             StringComparer.Ordinal)) {
                    content.Append('|').Append(relationship.RelationshipId).Append('=')
                        .Append(relationship.OpenXmlPart.Uri);
                }
            }

            var properties = document.PackageProperties;
            content.Append("|core|")
                .Append(properties.Creator).Append('|')
                .Append(properties.Title).Append('|')
                .Append(properties.Description).Append('|')
                .Append(properties.Category).Append('|')
                .Append(properties.Keywords).Append('|')
                .Append(properties.Subject).Append('|')
                .Append(properties.Revision).Append('|')
                .Append(properties.LastModifiedBy).Append('|')
                .Append(properties.Version).Append('|')
                .Append(properties.Created?.ToUniversalTime().Ticks).Append('|')
                .Append(properties.Modified?.ToUniversalTime().Ticks).Append('|')
                .Append(properties.LastPrinted?.ToUniversalTime().Ticks);

            using SHA256 sha = SHA256.Create();
            return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(content.ToString())));
        }

        private static void CollectParts(OpenXmlPart part, ISet<OpenXmlPart> parts) {
            if (!parts.Add(part)) return;
            foreach (IdPartPair child in part.Parts) CollectParts(child.OpenXmlPart, parts);
        }
    }
}
