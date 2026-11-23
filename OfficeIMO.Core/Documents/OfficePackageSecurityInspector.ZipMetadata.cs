using OfficeIMO.Core.Internal;
using System.IO.Compression;
using System.Xml;

namespace OfficeIMO {
    public static partial class OfficePackageSecurityInspector {
        private sealed class ZipXmlPart {
            internal ZipXmlPart(ZipArchiveEntry entry, string partName) {
                Entry = entry;
                PartName = partName;
            }

            internal ZipArchiveEntry Entry { get; }

            internal string PartName { get; }
        }

        private static void InspectContentTypes(
            ZipXmlPart contentTypesPart,
            ISet<string> packagePartNames,
            ZipPartIdentityIndex partIdentities,
            ISet<string> macroParts,
            ISet<string> embeddedParts,
            ISet<string> activeXParts,
            long maxXmlCharactersInPart,
            ICollection<OfficePackageSecurityFinding> findings) {
            var defaultContentTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try {
                using Stream stream = contentTypesPart.Entry.Open();
                using XmlReader reader = XmlReader.Create(
                    OfficeSmallXmlPartInput.Prepare(stream, contentTypesPart.Entry.Length),
                    CreateSecureXmlSettings(maxXmlCharactersInPart));
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    if (string.Equals(reader.LocalName, "Override", StringComparison.Ordinal)) {
                        string? partName = reader.GetAttribute("PartName");
                        string? contentType = reader.GetAttribute("ContentType");
                        string? normalizedPartName = NormalizeContentTypePartName(partName);
                        if (normalizedPartName != null) {
                            string identity = partIdentities.Resolve(normalizedPartName)
                                ?? contentTypesPart.PartName + "#" + normalizedPartName;
                            AddContentTypeClassification(contentType, identity, macroParts, embeddedParts, activeXParts);
                        }
                    } else if (string.Equals(reader.LocalName, "Default", StringComparison.Ordinal)) {
                        string? extension = reader.GetAttribute("Extension");
                        string? contentType = reader.GetAttribute("ContentType");
                        if (!string.IsNullOrWhiteSpace(extension) && !string.IsNullOrWhiteSpace(contentType)) {
                            defaultContentTypes[extension!.TrimStart('.')] = contentType!;
                        }
                    }
                }

                foreach (string partName in packagePartNames) {
                    int fileNameStart = partName.LastIndexOf('/');
                    int extensionStart = partName.LastIndexOf('.');
                    if (extensionStart <= fileNameStart || extensionStart == partName.Length - 1) continue;
                    string extension = partName.Substring(extensionStart + 1);
                    if (defaultContentTypes.TryGetValue(extension, out string? contentType)) {
                        AddContentTypeClassification(contentType, partName, macroParts, embeddedParts, activeXParts);
                    }
                }
            } catch (Exception exception) when (exception is XmlException || exception is InvalidDataException
                || exception is IOException) {
                findings.Add(Error(OfficePackageSecurityRule.MalformedPackage,
                    $"Content-types part '{contentTypesPart.PartName}' could not be parsed safely. {exception.Message}",
                    contentTypesPart.PartName));
            }
        }

        private static XmlReaderSettings CreateSecureXmlSettings(long maxXmlCharactersInPart) => new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            MaxCharactersInDocument = maxXmlCharactersInPart
        };

        private static void AddContentTypeClassification(
            string? contentType,
            string partName,
            ISet<string> macroParts,
            ISet<string> embeddedParts,
            ISet<string> activeXParts) {
            if (string.IsNullOrWhiteSpace(contentType)) return;
            if (ContainsIgnoreCase(contentType!, "vbaproject") || ContainsIgnoreCase(contentType!, "vbadata")) {
                macroParts.Add(partName);
            }
            if (ContainsIgnoreCase(contentType!, "oleobject")) embeddedParts.Add(partName);
            if (ContainsIgnoreCase(contentType!, "activex")) activeXParts.Add(partName);
        }

        private static void AddRelationshipClassification(
            string? relationshipType,
            string targetPart,
            ISet<string> macroParts,
            ISet<string> embeddedParts,
            ISet<string> activeXParts) {
            if (string.IsNullOrWhiteSpace(relationshipType)) return;
            if (ContainsIgnoreCase(relationshipType!, "/vbaproject")
                || ContainsIgnoreCase(relationshipType!, "/vbadata")) {
                macroParts.Add(targetPart);
            }
            if (ContainsIgnoreCase(relationshipType!, "/oleobject")
                || relationshipType!.EndsWith("/package", StringComparison.OrdinalIgnoreCase)) {
                embeddedParts.Add(targetPart);
            }
            if (ContainsIgnoreCase(relationshipType!, "/activex")
                || relationshipType!.EndsWith("/control", StringComparison.OrdinalIgnoreCase)) {
                activeXParts.Add(targetPart);
            }
        }

        private static string? NormalizeContentTypePartName(string? partName) {
            if (string.IsNullOrWhiteSpace(partName)) return null;
            string normalized = partName!.Replace('\\', '/');
            return normalized[0] == '/' ? normalized : "/" + normalized;
        }

        private static bool ContainsIgnoreCase(string value, string fragment) =>
            value.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0;
    }
}
