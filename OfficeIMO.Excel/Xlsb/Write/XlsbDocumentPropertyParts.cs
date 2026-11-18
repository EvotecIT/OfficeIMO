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
        /// Replaces the property parts of an existing XLSB package with the model's current values.
        /// Core and extended properties follow the package-property edits; custom properties follow their own edits.
        /// </summary>
        internal static void ApplyToRewrite(
            ExcelDocument document,
            Func<string, byte[]> readPart,
            IReadOnlyCollection<string> existingEntries,
            IDictionary<string, byte[]> replacements,
            ISet<string> deletions) {
            bool coreChanged = document.HasPackagePropertiesDirty;
            bool customChanged = document.HasCustomDocumentPropertiesDirty;
            if (!coreChanged && !customChanged) return;

            XlsbDocumentPropertyParts parts = Capture(document);
            XDocument relationships = LoadXml(replacements.TryGetValue("_rels/.rels", out byte[]? currentRelationships) ? currentRelationships : readPart("_rels/.rels"));
            XDocument types = LoadXml(replacements.TryGetValue("[Content_Types].xml", out byte[]? currentTypes) ? currentTypes : readPart("[Content_Types].xml"));
            var entries = new HashSet<string>(existingEntries, StringComparer.OrdinalIgnoreCase);
            if (coreChanged) {
                Apply(relationships, types, entries, replacements, deletions, CoreRelationshipType, CoreContentType, "docProps/core.xml", parts.Core, removeWhenEmpty: true);
                // An extended part the model could not read stays as it was rather than being dropped.
                if (parts.Extended != null) {
                    Apply(relationships, types, entries, replacements, deletions, ExtendedRelationshipType, ExtendedContentType, "docProps/app.xml", parts.Extended, removeWhenEmpty: false);
                }
            }
            if (customChanged) {
                Apply(relationships, types, entries, replacements, deletions, CustomRelationshipType, CustomContentType, "docProps/custom.xml", parts.Custom, removeWhenEmpty: true);
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
            bool removeWhenEmpty) {
            XNamespace rels = RelationshipsNamespace;
            XElement? relationship = relationships.Root!.Elements(rels + "Relationship")
                .FirstOrDefault(element => (string?)element.Attribute("Type") == relationshipType
                    && !string.Equals((string?)element.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase));
            string? target = (string?)relationship?.Attribute("Target");
            string? partName = string.IsNullOrWhiteSpace(target) ? null : XlsbPackagePartReader.ResolveTarget(string.Empty, target!);

            if (bytes == null) {
                if (!removeWhenEmpty || relationship == null) return;
                relationship.Remove();
                if (partName != null && entries.Contains(partName)) {
                    replacements.Remove(partName);
                    deletions.Add(partName);
                    RemoveOverride(types, partName);
                }
                return;
            }

            if (partName == null) {
                partName = defaultPartName;
                if (entries.Contains(partName) || replacements.ContainsKey(partName)) {
                    throw new NotSupportedException($"Native XLSB rewriting cannot add document properties because '{partName}' is already used by another package part.");
                }
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

        private static void RemoveOverride(XDocument types, string partName) {
            types.Root!.Elements()
                .Where(element => element.Name.LocalName == "Override"
                    && string.Equals(((string?)element.Attribute("PartName"))?.TrimStart('/'), partName, StringComparison.OrdinalIgnoreCase))
                .Remove();
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
            if (root == null || (!root.HasChildren && root is DocumentFormat.OpenXml.CustomProperties.Properties)) return null;
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
