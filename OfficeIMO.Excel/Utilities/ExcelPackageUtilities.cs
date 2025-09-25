using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel.Utilities {
    internal static class ExcelPackageUtilities {
        private const string ContentTypesEntry = "[Content_Types].xml";
        private const string WorkbookOverridePart = "/xl/workbook.xml";
        private const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        private const string AppPropsOverridePart = "/docProps/app.xml";
        private const string AppPropsContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        private const string CorePropsOverridePart = "/docProps/core.xml";
        private const string CorePropsContentType = "application/vnd.openxmlformats-package.core-properties+xml";

        internal static bool NormalizeContentTypes(string packagePath) {
            if (string.IsNullOrWhiteSpace(packagePath) || !File.Exists(packagePath)) {
                return false;
            }

            using var stream = new FileStream(packagePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return NormalizeContentTypes(stream, leaveOpen: true);
        }

        internal static bool NormalizeContentTypes(Stream packageStream, bool leaveOpen = false) {
            if (packageStream == null || !packageStream.CanRead || !packageStream.CanSeek) return false;
            using var archive = new ZipArchive(packageStream, ZipArchiveMode.Update, leaveOpen: leaveOpen);
            var entry = archive.GetEntry(ContentTypesEntry);
            if (entry == null) {
                return false;
            }

            string xml;
            using (var entryStream = entry.Open())
            using (var reader = new StreamReader(entryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: false)) {
                xml = reader.ReadToEnd();
            }

            if (string.IsNullOrWhiteSpace(xml)) {
                return false;
            }

            XDocument document;
            try {
                document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            } catch {
                document = XDocument.Parse(xml, LoadOptions.None);
            }

            var root = document.Root;
            if (root == null) {
                return false;
            }

            XNamespace ns = root.Name.Namespace;
            bool changed = false;

            var xmlDefaults = root.Elements(ns + "Default")
                .Where(e => string.Equals((string?)e.Attribute("Extension"), "xml", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (xmlDefaults.Count == 0) {
                root.AddFirst(new XElement(ns + "Default",
                    new XAttribute("Extension", "xml"),
                    new XAttribute("ContentType", "application/xml")));
                changed = true;
            } else {
                var first = xmlDefaults[0];
                if (!string.Equals((string?)first.Attribute("ContentType"), "application/xml", StringComparison.OrdinalIgnoreCase)) {
                    first.SetAttributeValue("ContentType", "application/xml");
                    changed = true;
                }
                for (int i = 1; i < xmlDefaults.Count; i++) {
                    xmlDefaults[i].Remove();
                    changed = true;
                }
            }

            bool EnsureOverride(string partName, string contentType) {
                var ov = root.Elements(ns + "Override")
                    .FirstOrDefault(e => string.Equals((string?)e.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase));
                if (ov == null) {
                    var newOverride = new XElement(ns + "Override",
                        new XAttribute("PartName", partName),
                        new XAttribute("ContentType", contentType));
                    var firstOverride = root.Elements(ns + "Override").FirstOrDefault();
                    if (firstOverride != null) firstOverride.AddBeforeSelf(newOverride); else root.Add(newOverride);
                    return true;
                }
                if (!string.Equals((string?)ov.Attribute("ContentType"), contentType, StringComparison.OrdinalIgnoreCase)) {
                    ov.SetAttributeValue("ContentType", contentType);
                    return true;
                }
                return false;
            }

            if (EnsureOverride(WorkbookOverridePart, WorkbookContentType)) changed = true;
            if (EnsureOverride(AppPropsOverridePart, AppPropsContentType)) changed = true;
            if (EnsureOverride(CorePropsOverridePart, CorePropsContentType)) changed = true;

            if (!changed) {
                return false;
            }

            document.Declaration = new XDeclaration("1.0", "utf-8", null);

            // Replace the entry content by recreating the entry
            entry.Delete();
            var newEntry = archive.CreateEntry(ContentTypesEntry, CompressionLevel.Optimal);
            using var output = newEntry.Open();
            var settings = new XmlWriterSettings {
                Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                Indent = false,
                OmitXmlDeclaration = false,
                NewLineHandling = NewLineHandling.None
            };
            using var writer = XmlWriter.Create(output, settings);
            document.Save(writer);
            writer.Flush();
            return true;
        }

        internal static ContentTypesSummary GetContentTypesSummary(string packagePath) {
            if (string.IsNullOrWhiteSpace(packagePath) || !File.Exists(packagePath)) {
                return new ContentTypesSummary(false, null, 0, false, null);
            }

            using var archive = ZipFile.OpenRead(packagePath);
            var entry = archive.GetEntry(ContentTypesEntry);
            if (entry == null) {
                return new ContentTypesSummary(false, null, 0, false, null);
            }

            string xml;
            using (var entryStream = entry.Open())
            using (var reader = new StreamReader(entryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: false)) {
                xml = reader.ReadToEnd();
            }

            if (string.IsNullOrWhiteSpace(xml)) {
                return new ContentTypesSummary(false, null, 0, false, null);
            }

            XDocument document;
            try {
                document = XDocument.Parse(xml, LoadOptions.None);
            } catch {
                return new ContentTypesSummary(false, null, 0, false, null);
            }

            var root = document.Root;
            if (root == null) {
                return new ContentTypesSummary(false, null, 0, false, null);
            }

            XNamespace ns = root.Name.Namespace;
            var xmlDefaults = root.Elements(ns + "Default")
                .Where(e => string.Equals((string?)e.Attribute("Extension"), "xml", StringComparison.OrdinalIgnoreCase))
                .ToList();
            string? xmlContentType = xmlDefaults.Count > 0 ? (string?)xmlDefaults[0].Attribute("ContentType") : null;
            var workbookOverride = root.Elements(ns + "Override")
                .FirstOrDefault(e => string.Equals((string?)e.Attribute("PartName"), WorkbookOverridePart, StringComparison.OrdinalIgnoreCase));
            string? workbookContentType = workbookOverride?.Attribute("ContentType")?.Value;

            return new ContentTypesSummary(
                xmlDefaults.Count > 0,
                xmlContentType,
                xmlDefaults.Count,
                workbookOverride != null,
                workbookContentType);
        }
    }

    internal readonly struct ContentTypesSummary {
        public bool HasXmlDefault { get; }
        public string? XmlDefaultContentType { get; }
        public int XmlDefaultCount { get; }
        public bool HasWorkbookOverride { get; }
        public string? WorkbookContentType { get; }

        public ContentTypesSummary(
            bool hasXmlDefault,
            string? xmlDefaultContentType,
            int xmlDefaultCount,
            bool hasWorkbookOverride,
            string? workbookContentType) {
            HasXmlDefault = hasXmlDefault;
            XmlDefaultContentType = xmlDefaultContentType;
            XmlDefaultCount = xmlDefaultCount;
            HasWorkbookOverride = hasWorkbookOverride;
            WorkbookContentType = workbookContentType;
        }
    }
}
