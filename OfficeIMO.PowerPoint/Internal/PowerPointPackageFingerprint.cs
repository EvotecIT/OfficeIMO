using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.PowerPoint {
    /// <summary>Creates deterministic fingerprints over a presentation package and its relationships.</summary>
    internal static class PowerPointPackageFingerprint {
        internal static string Create(PresentationDocument document,
            Action<OpenXmlPart, OpenXmlElement>? normalizeRoot = null,
            Func<OpenXmlPart, bool>? includePart = null,
            Func<OpenXmlPart, IdPartPair, bool>? includeRelationship = null,
            Func<OpenXmlPart, ReferenceRelationship, bool>? includeReferenceRelationship = null,
            Func<OpenXmlPart, DataPartReferenceRelationship, bool>?
                includeDataPartReferenceRelationship = null,
            bool includePackageProperties = true) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            var parts = new HashSet<OpenXmlPart>();
            foreach (IdPartPair pair in document.Parts) CollectParts(pair.OpenXmlPart, parts);

            var content = new StringBuilder();
            foreach (OpenXmlPart part in parts.OrderBy(item => item.Uri.ToString(), StringComparer.Ordinal)) {
                if (includePart != null && !includePart(part)) continue;
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
                    if (includeRelationship != null && !includeRelationship(part, relationship)) continue;
                    content.Append('|').Append(relationship.RelationshipId).Append('=')
                        .Append(relationship.OpenXmlPart.Uri);
                }
                foreach (DataPartReferenceRelationship relationship in part
                             .DataPartReferenceRelationships
                             .OrderBy(item => item.Id,
                                 StringComparer.Ordinal)) {
                    if (includeDataPartReferenceRelationship != null
                        && !includeDataPartReferenceRelationship(part,
                            relationship)) continue;
                    DataPart dataPart = relationship.DataPart;
                    content.Append("|data:").Append(relationship.Id)
                        .Append('=').Append(relationship.RelationshipType)
                        .Append('>').Append(dataPart.Uri).Append('|')
                        .Append(dataPart.ContentType).Append('|');
                    try {
                        using Stream stream = dataPart.GetStream(
                            FileMode.Open, FileAccess.Read);
                        using SHA256 dataHash = SHA256.Create();
                        content.Append(Convert.ToBase64String(
                            dataHash.ComputeHash(stream)));
                    } catch (Exception exception) when (
                        exception is InvalidDataException
                        || exception is IOException) {
                        content.Append("unreadable");
                    }
                }
                IEnumerable<ReferenceRelationship> references = part.ExternalRelationships
                    .Cast<ReferenceRelationship>()
                    .Concat(part.HyperlinkRelationships)
                    .OrderBy(item => item.Id, StringComparer.Ordinal);
                foreach (ReferenceRelationship relationship in references) {
                    if (includeReferenceRelationship != null
                        && !includeReferenceRelationship(part, relationship)) continue;
                    content.Append("|ref:").Append(relationship.Id).Append('=')
                        .Append(relationship.RelationshipType).Append('>')
                        .Append(relationship.Uri.OriginalString);
                }
            }

            if (includePackageProperties) {
                var properties = document.PackageProperties;
                content.Append("|core|");
                AppendPackageProperty(content, "Creator", properties.Creator);
                AppendPackageProperty(content, "Title", properties.Title);
                AppendPackageProperty(content, "Description", properties.Description);
                AppendPackageProperty(content, "Category", properties.Category);
                AppendPackageProperty(content, "ContentStatus", properties.ContentStatus);
                AppendPackageProperty(content, "ContentType", properties.ContentType);
                AppendPackageProperty(content, "Identifier", properties.Identifier);
                AppendPackageProperty(content, "Keywords", properties.Keywords);
                AppendPackageProperty(content, "Language", properties.Language);
                AppendPackageProperty(content, "Subject", properties.Subject);
                AppendPackageProperty(content, "Revision", properties.Revision);
                AppendPackageProperty(content, "LastModifiedBy", properties.LastModifiedBy);
                AppendPackageProperty(content, "Version", properties.Version);
                AppendPackageProperty(content, "Created", properties.Created?
                    .ToUniversalTime().Ticks.ToString(CultureInfo.InvariantCulture));
                AppendPackageProperty(content, "Modified", properties.Modified?
                    .ToUniversalTime().Ticks.ToString(CultureInfo.InvariantCulture));
                AppendPackageProperty(content, "LastPrinted", properties.LastPrinted?
                    .ToUniversalTime().Ticks.ToString(CultureInfo.InvariantCulture));
            }

            using SHA256 sha = SHA256.Create();
            return Convert.ToBase64String(sha.ComputeHash(Encoding.UTF8.GetBytes(content.ToString())));
        }

        private static void CollectParts(OpenXmlPart part, ISet<OpenXmlPart> parts) {
            if (!parts.Add(part)) return;
            foreach (IdPartPair child in part.Parts) CollectParts(child.OpenXmlPart, parts);
        }

        private static void AppendPackageProperty(StringBuilder content,
            string name, string? value) {
            string resolved = value ?? string.Empty;
            content.Append(name.Length).Append(':').Append(name)
                .Append('=').Append(resolved.Length).Append(':')
                .Append(resolved).Append(';');
        }
    }
}
