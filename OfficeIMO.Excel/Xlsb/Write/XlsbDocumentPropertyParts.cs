using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Serializes the workbook's core, extended and custom document properties as the ordinary
    /// OPC XML parts an XLSB package shares with XLSX (MS-XLSB 2.1.7: docProps/core.xml, app.xml, custom.xml).
    /// </summary>
    internal sealed class XlsbDocumentPropertyParts {
        internal const string CoreRelationshipType = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
        internal const string ExtendedRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
        internal const string CustomRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
        internal const string CoreContentType = "application/vnd.openxmlformats-package.core-properties+xml";
        internal const string ExtendedContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        internal const string CustomContentType = "application/vnd.openxmlformats-officedocument.custom-properties+xml";

        private const string RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string ContentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";

        private XlsbDocumentPropertyParts(byte[]? core, byte[]? extended, byte[]? custom) {
            Core = core;
            Extended = extended;
            Custom = custom;
        }

        internal byte[]? Core { get; }

        internal byte[]? Extended { get; }

        internal byte[]? Custom { get; }

        private IEnumerable<(string PartName, string RelationshipType, string ContentType, byte[] Bytes)> Parts {
            get {
                if (Core != null) yield return ("docProps/core.xml", CoreRelationshipType, CoreContentType, Core);
                if (Extended != null) yield return ("docProps/app.xml", ExtendedRelationshipType, ExtendedContentType, Extended);
                if (Custom != null) yield return ("docProps/custom.xml", CustomRelationshipType, CustomContentType, Custom);
            }
        }

        /// <summary>Captures the property parts from the workbook model, committing pending custom-property edits first.</summary>
        internal static XlsbDocumentPropertyParts Capture(ExcelDocument document) {
            ArgumentNullException.ThrowIfNull(document);
            document.SaveCustomDocumentProperties();
            SpreadsheetDocument package = document._spreadSheetDocument;
            return new XlsbDocumentPropertyParts(
                HasCoreProperties(package) ? ToUtf8(ExcelDocument.CreateCorePropertiesXml(package)) : null,
                SerializeRoot(package.ExtendedFilePropertiesPart?.Properties),
                SerializeRoot(package.CustomFilePropertiesPart?.Properties));
        }

        internal void AppendContentTypeOverrides(StringBuilder builder) {
            foreach (var part in Parts) {
                builder.Append("<Override PartName=\"/").Append(part.PartName)
                    .Append("\" ContentType=\"").Append(part.ContentType).Append("\"/>");
            }
        }

        internal string CreateRootRelationships(string officeDocumentTarget) {
            var builder = new StringBuilder(512);
            builder.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            builder.Append("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"")
                .Append(officeDocumentTarget).Append("\"/>");
            int id = 2;
            foreach (var part in Parts) {
                builder.Append("<Relationship Id=\"rId").Append(id++.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    .Append("\" Type=\"").Append(part.RelationshipType)
                    .Append("\" Target=\"").Append(part.PartName).Append("\"/>");
            }
            builder.Append("</Relationships>");
            return builder.ToString();
        }

        internal void WriteEntries(Action<string, byte[]> writeEntry) {
            foreach (var part in Parts) writeEntry(part.PartName, part.Bytes);
        }

        /// <summary>
        /// The property state of a loaded XLSB package as the model holds it, read without committing
        /// pending edits. Comparing against it also catches edits made directly through the Open XML package.
        /// </summary>
        internal sealed class Snapshot {
            private Snapshot(string? core, string? extended, string? custom) {
                Core = core;
                Extended = extended;
                Custom = custom;
            }

            internal string? Core { get; }

            internal string? Extended { get; }

            internal string? Custom { get; }

            internal static Snapshot Take(ExcelDocument document) {
                SpreadsheetDocument package = document._spreadSheetDocument;
                return new Snapshot(
                    HasCoreProperties(package) ? ExcelDocument.CreateCorePropertiesXml(package) : null,
                    package.ExtendedFilePropertiesPart?.Properties?.OuterXml,
                    package.CustomFilePropertiesPart?.Properties?.OuterXml);
            }

            internal bool Matches(ExcelDocument document) {
                Snapshot current = Take(document);
                return current.Core == Core && current.Extended == Extended && current.Custom == Custom;
            }
        }

        /// <summary>
        /// Replaces the property parts of an existing XLSB package with the model's current values.
        /// A part is rewritten when its wrapper API marked it changed or when it no longer matches
        /// the state captured after the package was loaded or last written.
        /// </summary>
        internal static void ApplyToRewrite(
            ExcelDocument document,
            Snapshot loaded,
            Func<string, byte[]> readPart,
            IReadOnlyCollection<string> existingEntries,
            IDictionary<string, byte[]> replacements,
            ISet<string> deletions) {
            Snapshot current = Snapshot.Take(document);
            bool packagePropertiesDirty = document.HasPackagePropertiesDirty;
            bool customDirty = document.HasCustomDocumentPropertiesDirty;
            bool coreChanged = packagePropertiesDirty || current.Core != loaded.Core;
            bool extendedChanged = packagePropertiesDirty || current.Extended != loaded.Extended;
            bool customChanged = customDirty || current.Custom != loaded.Custom;
            if (!coreChanged && !extendedChanged && !customChanged) return;

            XlsbDocumentPropertyParts parts = Capture(document);
            // Edits stay pending until the rewritten package is committed, so a failed save can be retried.
            if (customDirty) document.MarkCustomDocumentPropertiesPending();
            XDocument relationships = LoadXml(replacements.TryGetValue("_rels/.rels", out byte[]? currentRelationships) ? currentRelationships : readPart("_rels/.rels"));
            XDocument types = LoadXml(replacements.TryGetValue("[Content_Types].xml", out byte[]? currentTypes) ? currentTypes : readPart("[Content_Types].xml"));
            // Entry names are compared the way RewritePackage writes them, so docProps\core.xml
            // occupies docProps/core.xml.
            var entries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string entry in existingEntries) {
                string? normalized = TryNormalizePartName(entry);
                if (normalized != null) entries.Add(normalized);
            }
            // A part the model never held (absent or unreadable in the source) is not deleted, and its
            // relationship is never trusted as a write target; one the model held and no longer has is dropped.
            if (coreChanged) {
                Apply(relationships, types, entries, replacements, deletions, CoreRelationshipType, CoreContentType, "docProps/core.xml", parts.Core, sourcePartReadable: loaded.Core != null);
            }
            if (extendedChanged) {
                Apply(relationships, types, entries, replacements, deletions, ExtendedRelationshipType, ExtendedContentType, "docProps/app.xml", parts.Extended, sourcePartReadable: loaded.Extended != null);
            }
            if (customChanged) {
                Apply(relationships, types, entries, replacements, deletions, CustomRelationshipType, CustomContentType, "docProps/custom.xml", parts.Custom, sourcePartReadable: loaded.Custom != null);
            }
            replacements["_rels/.rels"] = SaveXml(relationships);
            replacements["[Content_Types].xml"] = SaveXml(types);
        }

        private static void Apply(
            XDocument relationships,
            XDocument types,
            HashSet<string> entries,
            IDictionary<string, byte[]> replacements,
            ISet<string> deletions,
            string relationshipType,
            string contentType,
            string defaultPartName,
            byte[]? bytes,
            bool sourcePartReadable) {
            XNamespace rels = RelationshipsNamespace;
            XElement? relationship = relationships.Root!.Elements(rels + "Relationship")
                .FirstOrDefault(element => (string?)element.Attribute("Type") == relationshipType
                    && !string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase));
            string? target = TryResolveRootTarget((string?)relationship?.Attribute("Target"));
            // Only a target the reader imported as this property part may be overwritten; an unreadable
            // or mistyped relationship can point at any part, including the workbook itself.
            string? partName = sourcePartReadable ? target : null;

            if (bytes == null) {
                if (!sourcePartReadable || relationship == null) return;
                relationship.Remove();
                if (partName != null && entries.Contains(partName)) {
                    replacements.Remove(partName);
                    deletions.Add(partName);
                    RemoveOverride(types, partName);
                }
                return;
            }

            if (partName == null) {
                partName = FindWritablePartName(relationships, types, entries, replacements, relationship, relationshipType, contentType, defaultPartName);
                relationship?.Remove();
                relationships.Root!.Add(new XElement(rels + "Relationship",
                    new XAttribute("Id", NextRelationshipId(relationships)),
                    new XAttribute("Type", relationshipType),
                    new XAttribute("Target", partName)));
            }

            deletions.Remove(partName);
            replacements[partName] = bytes;
            RemoveOverride(types, partName);
            XNamespace ct = ContentTypesNamespace;
            types.Root!.Add(new XElement(ct + "Override",
                new XAttribute("PartName", "/" + partName),
                new XAttribute("ContentType", contentType)));
        }

        /// <summary>
        /// Picks the part that receives a property kind the reader did not import. An occupied default
        /// part is replaced only when it already belongs to that kind (its declared content type matches
        /// and no other relationship claims it), which covers malformed XML of the right kind. A default
        /// name held by another part, such as different properties behind a mistyped relationship, is
        /// left alone and the properties go to a free name beside it.
        /// </summary>
        private static string FindWritablePartName(
            XDocument relationships,
            XDocument types,
            HashSet<string> entries,
            IDictionary<string, byte[]> replacements,
            XElement? relationship,
            string relationshipType,
            string contentType,
            string defaultPartName) {
            string stem = defaultPartName.Substring(0, defaultPartName.Length - ".xml".Length);
            string partName = defaultPartName;
            for (int suffix = 2; ; suffix++) {
                bool occupied = entries.Contains(partName) || replacements.ContainsKey(partName);
                if (!occupied || IsOwnedBy(relationships, types, replacements, relationship, relationshipType, contentType, partName)) return partName;
                partName = stem + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".xml";
            }
        }

        private static bool IsOwnedBy(
            XDocument relationships,
            XDocument types,
            IDictionary<string, byte[]> replacements,
            XElement? relationship,
            string relationshipType,
            string contentType,
            string partName) {
            // A part written earlier in this save belongs to the property kind that wrote it.
            if (replacements.ContainsKey(partName)) return false;
            if (!string.Equals(GetDeclaredContentType(types, partName), contentType, StringComparison.Ordinal)) return false;
            XNamespace rels = RelationshipsNamespace;
            return !relationships.Root!.Elements(rels + "Relationship").Any(element => element != relationship
                && !string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)
                && (string?)element.Attribute("Type") != relationshipType
                && string.Equals(TryResolveRootTarget((string?)element.Attribute("Target")), partName, StringComparison.OrdinalIgnoreCase));
        }

        private static string? GetDeclaredContentType(XDocument types, string partName) {
            XElement? partOverride = types.Root!.Elements().FirstOrDefault(element => element.Name.LocalName == "Override"
                && string.Equals(TryNormalizePartName((string?)element.Attribute("PartName")), partName, StringComparison.OrdinalIgnoreCase));
            if (partOverride != null) return (string?)partOverride.Attribute("ContentType");
            string extension = Path.GetExtension(partName).TrimStart('.');
            return (string?)types.Root.Elements().FirstOrDefault(element => element.Name.LocalName == "Default"
                && string.Equals((string?)element.Attribute("Extension"), extension, StringComparison.OrdinalIgnoreCase))?.Attribute("ContentType");
        }

        private static string? TryResolveRootTarget(string? target) {
            if (string.IsNullOrWhiteSpace(target)) return null;
            try {
                return XlsbPackagePartReader.ResolveTarget(string.Empty, target!);
            } catch (InvalidDataException) {
                return null;
            }
        }

        private static void RemoveOverride(XDocument types, string partName) {
            types.Root!.Elements()
                .Where(element => element.Name.LocalName == "Override"
                    && string.Equals(TryNormalizePartName((string?)element.Attribute("PartName")), partName, StringComparison.OrdinalIgnoreCase))
                .Remove();
        }

        private static string? TryNormalizePartName(string? partName) {
            // Directory entries and unsafe names are not parts; RewritePackage skips or rejects them itself.
            if (string.IsNullOrWhiteSpace(partName) || partName!.EndsWith("/", StringComparison.Ordinal) || partName.EndsWith("\\", StringComparison.Ordinal)) return null;
            try {
                return XlsbNativePackageWriter.NormalizePartName(partName);
            } catch (InvalidDataException) {
                return null;
            }
        }

        private static string NextRelationshipId(XDocument relationships) {
            var ids = new HashSet<string>(relationships.Root!.Elements().Select(element => (string?)element.Attribute("Id") ?? string.Empty), StringComparer.Ordinal);
            int next = 1;
            while (ids.Contains("rId" + next.ToString(System.Globalization.CultureInfo.InvariantCulture))) next++;
            return "rId" + next.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static bool HasCoreProperties(SpreadsheetDocument package) {
            var properties = package.PackageProperties;
            return !string.IsNullOrEmpty(properties.Title) || !string.IsNullOrEmpty(properties.Subject)
                || !string.IsNullOrEmpty(properties.Creator) || !string.IsNullOrEmpty(properties.Keywords)
                || !string.IsNullOrEmpty(properties.Description) || !string.IsNullOrEmpty(properties.LastModifiedBy)
                || !string.IsNullOrEmpty(properties.Revision) || !string.IsNullOrEmpty(properties.Category)
                || !string.IsNullOrEmpty(properties.Version) || !string.IsNullOrEmpty(properties.ContentStatus)
                || !string.IsNullOrEmpty(properties.Identifier) || !string.IsNullOrEmpty(properties.Language)
                || properties.LastPrinted.HasValue || properties.Created.HasValue || properties.Modified.HasValue;
        }

        private static byte[]? SerializeRoot(OpenXmlPartRootElement? root) {
            // An empty custom root is dropped, matching SaveCustomDocumentProperties; one that still
            // carries attributes keeps its metadata.
            if (root == null || (root is DocumentFormat.OpenXml.CustomProperties.Properties
                && !root.HasChildren && root.GetAttributes().Count == 0)) return null;
            using var stream = new MemoryStream();
            using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings {
                Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                CloseOutput = false
            })) {
                writer.WriteStartDocument(standalone: true);
                root.WriteTo(writer);
            }
            return stream.ToArray();
        }

        private static byte[] ToUtf8(string xml) => new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(xml);

        private static XDocument LoadXml(byte[] bytes) {
            using var stream = new MemoryStream(bytes, writable: false);
            using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null });
            return XDocument.Load(reader, LoadOptions.None);
        }

        private static byte[] SaveXml(XDocument document) {
            using var stream = new MemoryStream();
            using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings {
                Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                CloseOutput = false
            })) {
                document.Save(writer);
            }
            return stream.ToArray();
        }
    }
}
