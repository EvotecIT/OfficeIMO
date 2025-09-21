using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Utilities for discovering and extracting assets (masters) from existing VSDX files.
    /// </summary>
    public static class VisioAssets {
        private const string V = "http://schemas.microsoft.com/office/visio/2012/main";
        private const string REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string REL_ODC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        /// <summary>
        /// Lightweight info about a master available inside a VSDX file.
        /// </summary>
        public sealed class MasterInfo {
            public string Id { get; set; } = string.Empty;
            public string NameU { get; set; } = string.Empty;
            public string? Name { get; set; }
            public string RelationshipId { get; set; } = string.Empty;
            public override string ToString() => $"{Id}:{NameU} ({Name})";
        }

        /// <summary>
        /// Full master payload (MasterContents plus the masters.xml element) for importing.
        /// </summary>
        public sealed class MasterContent {
            public string Id { get; set; } = string.Empty;
            public string NameU { get; set; } = string.Empty;
            public XDocument MasterXml { get; set; } = new XDocument();
            public XElement MasterElement { get; set; } = new XElement(XName.Get("Master", V));
        }

        /// <summary>
        /// Lists masters available in the provided VSDX file.
        /// </summary>
        public static IReadOnlyList<MasterInfo> ListMasters(string vsdxPath) {
            using ZipArchive zip = ZipFile.OpenRead(vsdxPath);
            var mastersEntry = zip.GetEntry("visio/masters/masters.xml");
            if (mastersEntry == null) return Array.Empty<MasterInfo>();
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
            using var mStream = mastersEntry.Open();
            XDocument mastersDoc = XDocument.Load(mStream);
            XNamespace v = V;
            XNamespace r = REL_ODC;
            // map relId->target
            var relsEntry = zip.GetEntry("visio/masters/_rels/masters.xml.rels");
            var relMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (relsEntry != null) {
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
            foreach (var m in mastersDoc.Root!.Elements(v + "Master")) {
                string id = (string?)m.Attribute("ID") ?? string.Empty;
                string nameU = (string?)m.Attribute("NameU") ?? string.Empty;
                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(nameU)) continue;
                if (filter != null && !filter.Contains(nameU)) continue;
                string relId = (string?)m.Element(v + "Rel")?.Attribute(r + "id") ?? string.Empty;
                if (string.IsNullOrEmpty(relId) || !relMap.TryGetValue(relId, out string target)) continue;
                string partPath = "visio/masters/" + target.Replace("\\", "/");
                var part = zip.GetEntry(partPath);
                if (part == null) continue;
                using var pStream = part.Open();
                XDocument masterXml = XDocument.Load(pStream);
                result.Add(new MasterContent { Id = id, NameU = nameU, MasterXml = masterXml, MasterElement = new XElement(m) });
            }
            return result;
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
    }
}
