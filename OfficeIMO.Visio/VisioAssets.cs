using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Utilities for discovering and extracting assets (masters) from existing Visio packages.
    /// </summary>
    public static class VisioAssets {
        private const string V = "http://schemas.microsoft.com/office/visio/2012/main";
        private const string REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string REL_ODC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        /// <summary>Maximum uncompressed size accepted for a Visio master-related XML part.</summary>
        public const long MaxMasterXmlPartBytes = 12_000_000;
        /// <summary>Maximum uncompressed size accepted for a single embedded master relationship payload.</summary>
        public const long MaxMasterRelationshipBytes = 16_000_000;
        /// <summary>Maximum cumulative uncompressed size accepted for embedded master relationship payloads.</summary>
        public const long MaxTotalMasterRelationshipBytes = 64_000_000;
        /// <summary>Maximum number of master relationships copied from one Visio package import.</summary>
        public const int MaxMasterRelationships = 4096;

        /// <summary>
        /// Lightweight info about a master available inside a Visio package.
        /// </summary>
        public sealed class MasterInfo {
            /// <summary>Master identifier from masters.xml.</summary>
            public string Id { get; set; } = string.Empty;
            /// <summary>Universal (language-independent) master name.</summary>
            public string NameU { get; set; } = string.Empty;
            /// <summary>Localized master name, when present.</summary>
            public string? Name { get; set; }
            /// <summary>Relationship Id from masters.xml linking to the master part.</summary>
            public string RelationshipId { get; set; } = string.Empty;
            /// <summary>Returns a readable representation for diagnostics.</summary>
            public override string ToString() => $"{Id}:{NameU} ({Name})";
        }

        /// <summary>
        /// Full master payload (MasterContents plus the masters.xml element) for importing.
        /// </summary>
        public sealed class MasterContent {
            /// <summary>Master identifier.</summary>
            public string Id { get; set; } = string.Empty;
            /// <summary>Universal (language-independent) master name.</summary>
            public string NameU { get; set; } = string.Empty;
            /// <summary>The full XML of the master part (masterNNNN.xml).</summary>
            public XDocument MasterXml { get; set; } = new XDocument();
            /// <summary>The corresponding <c>Master</c> element from masters.xml.</summary>
            public XElement MasterElement { get; set; } = new XElement(XName.Get("Master", V));
            /// <summary>Relationships owned by the master part, including copied media payloads.</summary>
            public IList<MasterRelationshipContent> Relationships { get; } = new List<MasterRelationshipContent>();
        }

        /// <summary>
        /// A relationship from a master part to another package part or external target.
        /// </summary>
        public sealed class MasterRelationshipContent {
            /// <summary>Relationship id from the source master part.</summary>
            public string Id { get; set; } = string.Empty;
            /// <summary>Relationship type URI.</summary>
            public string Type { get; set; } = string.Empty;
            /// <summary>Original relationship target.</summary>
            public string Target { get; set; } = string.Empty;
            /// <summary>Whether the relationship points outside the package.</summary>
            public bool IsExternal { get; set; }
            /// <summary>Content type of the internal target part, when known.</summary>
            public string ContentType { get; set; } = string.Empty;
            /// <summary>File extension for the copied internal target part.</summary>
            public string Extension { get; set; } = string.Empty;
            /// <summary>Raw internal target part bytes.</summary>
            public byte[]? Data { get; set; }
        }

        /// <summary>
        /// Visual context from a Visio package document part that imported masters may inherit.
        /// </summary>
        public sealed class PackageVisualContext {
            /// <summary>Raw visio/document.xml from the source package.</summary>
            public XDocument? DocumentXml { get; set; }
            /// <summary>Raw theme XML from the source package, when present.</summary>
            public XDocument? ThemeXml { get; set; }
        }

        /// <summary>
        /// Lists masters available in the provided Visio package.
        /// </summary>
        public static IReadOnlyList<MasterInfo> ListMasters(string vsdxPath) {
            using ZipArchive zip = ZipFile.OpenRead(vsdxPath);
            var mastersEntry = zip.GetEntry("visio/masters/masters.xml");
            if (mastersEntry == null) return Array.Empty<MasterInfo>();
            EnsureZipEntryWithinLimit(mastersEntry, MaxMasterXmlPartBytes, "masters XML part");
            using var s = mastersEntry.Open();
            XDocument doc = XDocument.Load(s);
            XNamespace v = V;
            return doc.Root!.Elements(v + "Master").Select(m => new MasterInfo {
                Id = (string?)m.Attribute("ID") ?? string.Empty,
                NameU = (string?)m.Attribute("NameU") ?? string.Empty,
                Name = (string?)m.Attribute("Name"),
                RelationshipId = (string?)m.Element(v + "Rel")?.Attribute(XName.Get("id", REL_ODC)) ?? string.Empty
            }).Where(x => !string.IsNullOrEmpty(x.Id) && !string.IsNullOrEmpty(x.NameU)).ToList();
        }

        /// <summary>
        /// Loads full master contents for selected masters. If <paramref name="filterNames"/> is null, loads all.
        /// </summary>
        public static IReadOnlyList<MasterContent> LoadMasterContents(string vsdxPath, IEnumerable<string>? filterNames = null) {
            HashSet<string>? filter = filterNames != null ? new HashSet<string>(filterNames, StringComparer.OrdinalIgnoreCase) : null;
            using ZipArchive zip = ZipFile.OpenRead(vsdxPath);
            var mastersEntry = zip.GetEntry("visio/masters/masters.xml");
            if (mastersEntry == null) return Array.Empty<MasterContent>();
            EnsureZipEntryWithinLimit(mastersEntry, MaxMasterXmlPartBytes, "masters XML part");
            using var mStream = mastersEntry.Open();
            XDocument mastersDoc = XDocument.Load(mStream);
            XNamespace v = V;
            XNamespace r = REL_ODC;
            // map relId->target
            var relsEntry = zip.GetEntry("visio/masters/_rels/masters.xml.rels");
            var relMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (relsEntry != null) {
                EnsureZipEntryWithinLimit(relsEntry, MaxMasterXmlPartBytes, "masters relationship part");
                using var relsStream = relsEntry.Open();
                var relsDoc = XDocument.Load(relsStream);
                XNamespace pr = REL_PKG;
                foreach (var e in relsDoc.Root!.Elements(pr + "Relationship")) {
                    string id = (string?)e.Attribute("Id") ?? string.Empty;
                    string target = (string?)e.Attribute("Target") ?? string.Empty;
                    if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target)) relMap[id] = target;
                }
            }
            List<MasterContent> result = new();
            long totalRelationshipBytes = 0;
            int totalRelationships = 0;
            foreach (var m in mastersDoc.Root!.Elements(v + "Master")) {
                string id = (string?)m.Attribute("ID") ?? string.Empty;
                string nameU = (string?)m.Attribute("NameU") ?? string.Empty;
                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(nameU)) continue;
                if (filter != null && !filter.Contains(nameU)) continue;
                string relId = (string?)m.Element(v + "Rel")?.Attribute(r + "id") ?? string.Empty;
                if (string.IsNullOrEmpty(relId) || !relMap.TryGetValue(relId, out var targetTmp)) continue;
                string target = targetTmp;
                string partPath = "visio/masters/" + target.Replace("\\", "/");
                var part = zip.GetEntry(partPath);
                if (part == null) continue;
                EnsureZipEntryWithinLimit(part, MaxMasterXmlPartBytes, "master XML part");
                using var pStream = part.Open();
                XDocument masterXml = XDocument.Load(pStream);
                MasterContent content = new() { Id = id, NameU = nameU, MasterXml = masterXml, MasterElement = new XElement(m) };
                foreach (MasterRelationshipContent relationship in LoadMasterRelationships(zip, partPath, ref totalRelationshipBytes, ref totalRelationships)) {
                    content.Relationships.Add(relationship);
                }

                result.Add(content);
            }
            return result;
        }

        /// <summary>
        /// Loads package-level visual context that imported masters can depend on for inherited colors, styles, and theme values.
        /// </summary>
        public static PackageVisualContext LoadVisualContext(string vsdxPath) {
            using ZipArchive zip = ZipFile.OpenRead(vsdxPath);
            PackageVisualContext context = new();
            ZipArchiveEntry? documentEntry = zip.GetEntry("visio/document.xml");
            if (documentEntry != null) {
                EnsureZipEntryWithinLimit(documentEntry, MaxMasterXmlPartBytes, "document XML part");
                using Stream stream = documentEntry.Open();
                context.DocumentXml = XDocument.Load(stream);
            }

            ZipArchiveEntry? themeEntry = zip.GetEntry("visio/theme/theme1.xml");
            if (themeEntry != null) {
                EnsureZipEntryWithinLimit(themeEntry, MaxMasterXmlPartBytes, "theme XML part");
                using Stream stream = themeEntry.Open();
                context.ThemeXml = XDocument.Load(stream);
            }

            return context;
        }

        /// <summary>
        /// Extracts masters from a VSDX file to a folder as standalone XML files (one per master).
        /// </summary>
        public static void ExtractMasters(string vsdxPath, string outputFolder, IEnumerable<string>? filterNames = null) {
            Directory.CreateDirectory(outputFolder);
            foreach (var m in LoadMasterContents(vsdxPath, filterNames)) {
                string safeName = string.Concat(m.NameU.Select(ch => char.IsLetterOrDigit(ch) ? ch : '_'));
                string filePath = Path.Combine(outputFolder, $"{m.Id}-{safeName}.xml");
                using var fs = File.Create(filePath);
                m.MasterXml.Save(fs, SaveOptions.DisableFormatting);
            }
        }

        private static IEnumerable<MasterRelationshipContent> LoadMasterRelationships(ZipArchive zip, string masterPartPath, ref long totalRelationshipBytes, ref int totalRelationships) {
            string fileName = Path.GetFileName(masterPartPath.Replace('\\', '/'));
            string relsPath = "visio/masters/_rels/" + fileName + ".rels";
            ZipArchiveEntry? relsEntry = zip.GetEntry(relsPath);
            if (relsEntry == null) {
                return Array.Empty<MasterRelationshipContent>();
            }

            EnsureZipEntryWithinLimit(relsEntry, MaxMasterXmlPartBytes, "master relationship XML part");
            using Stream relsStream = relsEntry.Open();
            XDocument relsDoc = XDocument.Load(relsStream);
            XNamespace pr = REL_PKG;
            List<MasterRelationshipContent> relationships = new();
            foreach (XElement rel in relsDoc.Root?.Elements(pr + "Relationship") ?? Enumerable.Empty<XElement>()) {
                string id = (string?)rel.Attribute("Id") ?? string.Empty;
                string type = (string?)rel.Attribute("Type") ?? string.Empty;
                string target = (string?)rel.Attribute("Target") ?? string.Empty;
                bool external = string.Equals((string?)rel.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(type) || string.IsNullOrWhiteSpace(target)) {
                    continue;
                }

                totalRelationships++;
                if (totalRelationships > MaxMasterRelationships) {
                    throw new InvalidDataException($"Visio package master relationships exceed {MaxMasterRelationships} entries.");
                }

                MasterRelationshipContent relationship = new() {
                    Id = id,
                    Type = type,
                    Target = target,
                    IsExternal = external
                };

                if (!external) {
                    string masterDirectory = (Path.GetDirectoryName(masterPartPath.Replace('\\', '/')) ?? string.Empty).Replace('\\', '/');
                    string targetPath = ResolveZipPath(masterDirectory, target);
                    ZipArchiveEntry? targetEntry = zip.GetEntry(targetPath);
                    if (targetEntry == null) {
                        continue;
                    }

                    relationship.Data = ReadMasterRelationshipBytes(targetEntry, ref totalRelationshipBytes);
                    relationship.Extension = Path.GetExtension(targetPath);
                    relationship.ContentType = GuessContentType(relationship.Extension);
                }

                relationships.Add(relationship);
            }

            return relationships;
        }

        private static byte[] ReadMasterRelationshipBytes(ZipArchiveEntry entry, ref long totalRelationshipBytes) {
            EnsureZipEntryWithinLimit(entry, MaxMasterRelationshipBytes, "master relationship target");
            if (totalRelationshipBytes > MaxTotalMasterRelationshipBytes - entry.Length) {
                throw new InvalidDataException($"Visio package master relationship targets exceed {MaxTotalMasterRelationshipBytes} total bytes.");
            }

            using Stream source = entry.Open();
            using MemoryStream buffer = entry.Length > 0 && entry.Length <= int.MaxValue
                ? new MemoryStream((int)entry.Length)
                : new MemoryStream();
            CopyToWithLimit(source, buffer, MaxMasterRelationshipBytes, entry.FullName);
            if (totalRelationshipBytes > MaxTotalMasterRelationshipBytes - buffer.Length) {
                throw new InvalidDataException($"Visio package master relationship targets exceed {MaxTotalMasterRelationshipBytes} total bytes.");
            }

            totalRelationshipBytes += buffer.Length;
            return buffer.ToArray();
        }

        private static void EnsureZipEntryWithinLimit(ZipArchiveEntry entry, long maxBytes, string description) {
            if (entry.Length > maxBytes) {
                throw new InvalidDataException($"Visio package {description} '{entry.FullName}' exceeds {maxBytes} bytes.");
            }
        }

        private static void CopyToWithLimit(Stream source, Stream destination, long maxBytes, string entryName) {
            byte[] buffer = new byte[81920];
            long copied = 0;
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0) {
                copied += read;
                if (copied > maxBytes) {
                    throw new InvalidDataException($"Visio package master relationship target '{entryName}' exceeds {maxBytes} bytes.");
                }

                destination.Write(buffer, 0, read);
            }
        }

        private static string ResolveZipPath(string basePath, string target) {
            string normalizedTarget = target.Replace('\\', '/');
            if (normalizedTarget.StartsWith("/", StringComparison.Ordinal)) {
                return normalizedTarget.TrimStart('/');
            }

            string combined = string.IsNullOrWhiteSpace(basePath) ? normalizedTarget : basePath.TrimEnd('/') + "/" + normalizedTarget;
            Stack<string> parts = new();
            foreach (string part in combined.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries)) {
                if (part == ".") {
                    continue;
                }

                if (part == "..") {
                    if (parts.Count > 0) {
                        parts.Pop();
                    }
                    continue;
                }

                parts.Push(part);
            }

            return string.Join("/", parts.Reverse());
        }

        private static string GuessContentType(string extension) {
            return extension.TrimStart('.').ToLowerInvariant() switch {
                "emf" => "image/x-emf",
                "png" => "image/png",
                "jpg" or "jpeg" => "image/jpeg",
                "gif" => "image/gif",
                "svg" => "image/svg+xml",
                "tif" or "tiff" => "image/tiff",
                _ => "application/octet-stream"
            };
        }
    }
}
