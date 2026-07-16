using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace OfficeIMO.Drawing {
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
            ISet<string> macroParts,
            ISet<string> embeddedParts,
            ISet<string> activeXParts,
            ICollection<OfficePackageSecurityFinding> findings) {
            var defaultContentTypes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try {
                using Stream stream = contentTypesPart.Entry.Open();
                using XmlReader reader = XmlReader.Create(stream, CreateSecureXmlSettings(contentTypesPart.Entry.Length));
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    if (string.Equals(reader.LocalName, "Override", StringComparison.Ordinal)) {
                        string? partName = reader.GetAttribute("PartName");
                        string? contentType = reader.GetAttribute("ContentType");
                        string? normalizedPartName = NormalizeContentTypePartName(partName);
                        if (normalizedPartName != null) {
                            AddContentTypeClassification(contentType, normalizedPartName, macroParts, embeddedParts, activeXParts);
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

        private static XmlReaderSettings CreateSecureXmlSettings(long entryLength) => new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreWhitespace = true,
            MaxCharactersInDocument = Math.Max(1024L, entryLength + 1L)
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

        private static string? ResolveRelationshipTarget(string relationshipPartName, string? target) {
            if (string.IsNullOrWhiteSpace(target)) return null;
            if (Uri.TryCreate(target, UriKind.Absolute, out Uri? absolute) && absolute.IsAbsoluteUri) return null;

            string normalizedRelationshipPart = relationshipPartName.Replace('\\', '/');
            int relationshipsSegment = normalizedRelationshipPart.LastIndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
            if (relationshipsSegment < 0) return null;

            string sourcePrefix = normalizedRelationshipPart.Substring(0, relationshipsSegment);
            string relationshipFile = normalizedRelationshipPart.Substring(relationshipsSegment + 7);
            string sourceDirectory;
            if (string.Equals(relationshipFile, ".rels", StringComparison.OrdinalIgnoreCase)) {
                sourceDirectory = "/";
            } else if (relationshipFile.EndsWith(RelationshipSuffix, StringComparison.OrdinalIgnoreCase)) {
                string sourceFile = relationshipFile.Substring(0, relationshipFile.Length - RelationshipSuffix.Length);
                string sourcePart = sourcePrefix + "/" + sourceFile;
                int separator = sourcePart.LastIndexOf('/');
                sourceDirectory = separator <= 0 ? "/" : sourcePart.Substring(0, separator + 1);
            } else {
                return null;
            }

            string normalizedTarget = target!.Replace('\\', '/');
            int suffix = normalizedTarget.IndexOfAny(new[] { '?', '#' });
            if (suffix >= 0) normalizedTarget = normalizedTarget.Substring(0, suffix);
            string combined = normalizedTarget.StartsWith("/", StringComparison.Ordinal)
                ? normalizedTarget
                : sourceDirectory + normalizedTarget;
            var segments = new List<string>();
            foreach (string segment in combined.Split('/')) {
                if (segment.Length == 0 || segment == ".") continue;
                if (segment == "..") {
                    if (segments.Count == 0) return null;
                    segments.RemoveAt(segments.Count - 1);
                } else {
                    segments.Add(segment);
                }
            }
            return "/" + string.Join("/", segments);
        }

        private static bool ContainsIgnoreCase(string value, string fragment) =>
            value.IndexOf(fragment, StringComparison.OrdinalIgnoreCase) >= 0;
    }
}
