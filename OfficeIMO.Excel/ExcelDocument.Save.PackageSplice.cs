using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
        private const string CorePropertiesContentType = "application/vnd.openxmlformats-package.core-properties+xml";

        /// <summary>
        /// Saves a loaded package by rewriting only the parts whose content differs from the package it was loaded from.
        /// </summary>
        /// <remarks>
        /// The live package must prove it has the source's structure: the same parts and content types, and the same
        /// relationships apart from parts it has dropped, such as the calculation chain. Part content is compared by
        /// size and CRC, because parts can change through streams that leave no trace on the document.
        /// </remarks>
        private bool TryWriteSplicedPackage(Stream destination, ExcelSaveOptions? options) {
            byte[]? package = TrySplicePackage(options);
            if (package == null) return false;

            PrepareDestinationStreamForWrite(destination);
            destination.Write(package, 0, package.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }
            MarkPackageClean(package);
            return true;
        }

        private bool TrySaveSplicedPackageToFile(string targetPath, ExcelSaveOptions? options, string? fallbackReason) {
            byte[]? package = TrySplicePackage(options);
            if (package == null) return false;

            CommitPreparedPackageToFile(targetPath, package);
            MarkPackageClean(package);
            FilePath = targetPath;
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard(fallbackReason);
            return true;
        }

        private byte[]? TrySplicePackage(ExcelSaveOptions? options) {
            byte[]? source = _packageBaselineBytes;
            if (source == null
                || options?.DisableFastPackageWriter == true
                || _spreadSheetDocument.DigitalSignatureOriginPart != null
                || _spreadSheetDocument.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null
                || CollectLoadedCarriageReturnParts(WorkbookPartRoot).Count > 0) {
                return null;
            }

            // Writing the package only to read its parts back here compresses every part, so that waits for a stream
            // that outlives this save.
            if (_packageStream != null || _usesFileBackedPackage) _spreadSheetDocument.Save();
            byte[]? package = ExcelPackageEntrySplicer.TryEdit(source, SpliceLivePackage);
            if (package == null) return null;

            ThrowIfPackageMaterializationExceedsLimit(package.LongLength, options);
            ThrowIfOpenXmlValidationFails(package, options);
            return package;
        }

        private void SpliceLivePackage(ExcelPackageEntrySplicer splicer) {
            var relationshipSets = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            var parts = new Dictionary<string, OpenXmlPart>(StringComparer.OrdinalIgnoreCase);
            CollectLiveRelationships(_spreadSheetDocument, string.Empty, relationshipSets, parts);

            KeyValuePair<string, byte[]> core = RewriteCoreProperties(splicer);
            var removedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var expectedEntries = new HashSet<string>(parts.Keys.Select(static name => name.TrimStart('/')), StringComparer.OrdinalIgnoreCase) {
                "[Content_Types].xml"
            };

            foreach (string relationships in splicer.Names.Where(static name => name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)).ToArray()) {
                string owner = GetRelationshipOwner(relationships);
                expectedEntries.Add(relationships);
                relationshipSets.TryGetValue(owner, out HashSet<string>? live);
                ReconcileRelationships(splicer, relationships, owner, live ?? new HashSet<string>(StringComparer.Ordinal), parts, removedParts);
                relationshipSets.Remove(owner);
            }
            if (relationshipSets.Values.Any(static set => set.Count > 0)) throw new InvalidDataException("A part gained relationships.");

            foreach (string name in splicer.Names.ToArray()) {
                if (expectedEntries.Contains(name)) continue;
                if (!removedParts.Contains(name) || splicer.Contains(GetRelationshipEntry(name))) throw new InvalidDataException($"Package entry '{name}' has no live part.");
                splicer.Remove(name);
            }
            ReconcileContentTypes(splicer, parts, removedParts);

            var buffer = new ExcelPartBuffer();
            try {
                foreach (KeyValuePair<string, OpenXmlPart> pair in parts) {
                    string name = pair.Key.TrimStart('/');
                    if (!splicer.Contains(name)) throw new InvalidDataException($"Part '{name}' is new.");
                    if (string.Equals(name, core.Key, StringComparison.OrdinalIgnoreCase)) continue;
                    SpliceLivePart(splicer, name, pair.Value, buffer);
                }
            } finally {
                buffer.Release();
            }
            splicer.Replace(core.Key, core.Value);
        }

        /// <summary>Leaves a changed root in its part as well, so nothing that reads the live package sees it stale.</summary>
        private static void SpliceLivePart(ExcelPackageEntrySplicer splicer, string name, OpenXmlPart part, ExcelPartBuffer buffer) {
            bool serialized = part.IsRootElementLoaded && part.RootElement != null;
            buffer.Clear();
            if (serialized) {
                if (part.RootElement is not Worksheet worksheet || !ExcelWorksheetGridWriter.TrySave(worksheet, buffer)) {
                    buffer.Clear();
                    part.RootElement!.Save(buffer);
                }
            } else {
                using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
                buffer.ReadFrom(stream);
            }
            if (splicer.Matches(name, buffer.Buffer, buffer.ContentLength)) return;
            var content = new byte[buffer.ContentLength];
            Buffer.BlockCopy(buffer.Buffer, 0, content, 0, content.Length);
            splicer.Replace(name, content);
            if (serialized) {
                using Stream target = part.GetStream(FileMode.Create, FileAccess.Write);
                target.Write(content, 0, content.Length);
            }
        }

        private static void CollectLiveRelationships(
            OpenXmlPartContainer container,
            string owner,
            Dictionary<string, HashSet<string>> relationshipSets,
            Dictionary<string, OpenXmlPart> parts) {
            if (container.DataPartReferenceRelationships.Any()) throw new InvalidDataException("Data parts are outside the splice model.");
            var relationships = new HashSet<string>(StringComparer.Ordinal);
            relationshipSets.Add(owner, relationships);
            foreach (IdPartPair pair in container.Parts) {
                OpenXmlPart part = pair.OpenXmlPart;
                string path = part.Uri.OriginalString;
                relationships.Add(pair.RelationshipId + "\n" + part.RelationshipType + "\n" + path);
                if (parts.TryGetValue(path, out OpenXmlPart? known)) {
                    if (!ReferenceEquals(known, part)) throw new InvalidDataException("Two parts share a name.");
                    continue;
                }
                parts.Add(path, part);
                CollectLiveRelationships(part, path, relationshipSets, parts);
            }
            foreach (ExternalRelationship relationship in container.ExternalRelationships) {
                relationships.Add(relationship.Id + "\n" + relationship.RelationshipType + "\nexternal:" + relationship.Uri.OriginalString);
            }
            foreach (HyperlinkRelationship relationship in container.HyperlinkRelationships) {
                relationships.Add(relationship.Id + "\n" + relationship.RelationshipType + "\n"
                    + (relationship.IsExternal ? "external:" : string.Empty) + relationship.Uri.OriginalString);
            }
        }

        private static void ReconcileRelationships(
            ExcelPackageEntrySplicer splicer,
            string entry,
            string owner,
            HashSet<string> live,
            Dictionary<string, OpenXmlPart> parts,
            HashSet<string> removedParts) {
            XDocument document = LoadPackageXml(splicer.Read(entry));
            XElement root = document.Root ?? throw new InvalidDataException("Relationships are empty.");
            bool changed = false;
            int matched = 0;
            foreach (XElement relationship in root.Elements(RelationshipsNamespace + "Relationship").ToArray()) {
                bool external = string.Equals((string?)relationship.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase);
                string target = (string?)relationship.Attribute("Target") ?? throw new InvalidDataException("A relationship has no target.");
                string resolved = external ? "external:" + target : ResolvePackagePath(owner, target);
                string key = (string?)relationship.Attribute("Id") + "\n" + (string?)relationship.Attribute("Type") + "\n" + resolved;
                if (live.Contains(key)) {
                    matched++;
                } else if (!external && !parts.ContainsKey(resolved)) {
                    removedParts.Add(resolved.TrimStart('/'));
                    relationship.Remove();
                    changed = true;
                } else {
                    throw new InvalidDataException("A relationship changed.");
                }
            }
            if (matched != live.Count) throw new InvalidDataException("A relationship was added.");
            if (changed) splicer.Replace(entry, WritePackageXml(document));
        }

        private static void ReconcileContentTypes(
            ExcelPackageEntrySplicer splicer,
            Dictionary<string, OpenXmlPart> parts,
            HashSet<string> removedParts) {
            XDocument document = LoadPackageXml(splicer.Read("[Content_Types].xml"));
            XElement root = document.Root ?? throw new InvalidDataException("Content types are empty.");
            var overrides = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement item in root.Elements(ContentTypesNamespace + "Override")) {
                overrides[(string?)item.Attribute("PartName") ?? string.Empty] = item;
            }
            var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement item in root.Elements(ContentTypesNamespace + "Default")) {
                defaults[(string?)item.Attribute("Extension") ?? string.Empty] = (string?)item.Attribute("ContentType") ?? string.Empty;
            }

            foreach (KeyValuePair<string, OpenXmlPart> pair in parts) {
                string extension = Path.GetExtension(pair.Key).TrimStart('.');
                string? declared = overrides.TryGetValue(pair.Key, out XElement? item)
                    ? (string?)item.Attribute("ContentType")
                    : defaults.TryGetValue(extension, out string? fallback) ? fallback : null;
                if (!string.Equals(declared, pair.Value.ContentType, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(declared, CorePropertiesContentType, StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidDataException($"The content type of '{pair.Key}' changed.");
                }
            }

            bool changed = false;
            foreach (string removed in removedParts) {
                if (overrides.TryGetValue("/" + removed, out XElement? item)) {
                    item.Remove();
                    changed = true;
                }
            }
            if (changed) splicer.Replace("[Content_Types].xml", WritePackageXml(document));
        }

        private static byte[] WritePackageXml(XDocument document) {
            using var output = new MemoryStream();
            using (var writer = System.Xml.XmlWriter.Create(output, new System.Xml.XmlWriterSettings { Encoding = new System.Text.UTF8Encoding(false) })) {
                document.Save(writer);
            }
            return output.ToArray();
        }

        private static string GetRelationshipOwner(string entry) {
            int marker = entry.LastIndexOf("_rels/", StringComparison.OrdinalIgnoreCase);
            if (marker < 0 || !entry.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Unexpected relationship entry.");
            string name = entry.Substring(marker + 6, entry.Length - marker - 11);
            return name.Length == 0 ? string.Empty : "/" + entry.Substring(0, marker) + name;
        }

        private static string GetRelationshipEntry(string entry) {
            int slash = entry.LastIndexOf('/');
            return entry.Substring(0, slash + 1) + "_rels/" + entry.Substring(slash + 1) + ".rels";
        }

        private static string ResolvePackagePath(string owner, string target) {
            string directory = owner.Length == 0 ? "/" : owner.Substring(0, owner.LastIndexOf('/') + 1);
            return Uri.UnescapeDataString(new Uri(new Uri("http://package" + directory), target).AbsolutePath);
        }
    }
}
