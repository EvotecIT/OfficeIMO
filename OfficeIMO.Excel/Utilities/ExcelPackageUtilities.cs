using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel.Utilities {
    internal static partial class ExcelPackageUtilities {
        private const string ContentTypesEntry = "[Content_Types].xml";
        internal const long MaximumContentTypesEntryBytes = 16L * 1024L * 1024L;
        private const string WorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
        private const string MacroEnabledWorkbookContentType = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
        private const string TemplateWorkbookContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
        private const string MacroEnabledTemplateWorkbookContentType = "application/vnd.ms-excel.template.macroEnabled.main+xml";
        internal const string AddInWorkbookContentType = "application/vnd.ms-excel.addin.macroEnabled.main+xml";
        private const string AppPropsContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
        private const string CorePropsContentType = "application/vnd.openxmlformats-package.core-properties+xml";

        internal static bool NormalizeContentTypes(string packagePath) {
            if (string.IsNullOrWhiteSpace(packagePath) || !File.Exists(packagePath)) {
                return false;
            }

            using var stream = new FileStream(packagePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return NormalizeContentTypes(stream, leaveOpen: true);
        }

        internal static bool NormalizeContentTypes(
            Stream packageStream,
            bool leaveOpen = false,
            long maxCharactersInPart = 0) {
            if (packageStream == null || !packageStream.CanRead || !packageStream.CanSeek) return false;
            long originalPosition = packageStream.Position;
            if (!NeedsContentTypeNormalization(packageStream, maxCharactersInPart)) {
                packageStream.Position = originalPosition;
                return false;
            }

            packageStream.Position = originalPosition;
            using var archive = new ZipArchive(packageStream, ZipArchiveMode.Update, leaveOpen: leaveOpen);
            byte[]? normalized = BuildNormalizedContentTypes(archive, maxCharactersInPart);
            if (normalized == null) return false;

            archive.GetEntry(ContentTypesEntry)!.Delete();
            var newEntry = archive.CreateEntry(ContentTypesEntry, CompressionLevel.Optimal);
            using var output = newEntry.Open();
            output.Write(normalized, 0, normalized.Length);
            return true;
        }

        /// <summary>
        /// Normalizes a package held in memory by replacing only its content-types entry. Returns false when the
        /// package is outside the plain ZIP shape the splicer handles; the stream route then applies.
        /// </summary>
        internal static bool TryNormalizeContentTypes(byte[] package, long maxCharactersInPart, out byte[] normalized, out bool changed) {
            normalized = package;
            changed = false;
            using var input = new MemoryStream(package, writable: false);
            if (!NeedsContentTypeNormalization(input, maxCharactersInPart)) return true;

            input.Position = 0;
            using var archive = new ZipArchive(input, ZipArchiveMode.Read, leaveOpen: true);
            byte[]? contentTypes = BuildNormalizedContentTypes(archive, maxCharactersInPart);
            if (contentTypes == null) return true;

            byte[]? spliced = ExcelPackageEntrySplicer.TryEdit(package, splicer => splicer.Replace(ContentTypesEntry, contentTypes));
            if (spliced == null) return false;
            normalized = spliced;
            changed = true;
            return true;
        }

        /// <summary>Returns the corrected content-types part, or null when the package has none to correct.</summary>
        private static byte[]? BuildNormalizedContentTypes(ZipArchive archive, long maxCharactersInPart) {
            var entry = archive.GetEntry(ContentTypesEntry);
            if (entry == null) {
                return null;
            }

            string xml = ReadContentTypesXml(entry, maxCharactersInPart);

            if (string.IsNullOrWhiteSpace(xml)) {
                return null;
            }

            XDocument document;
            try {
                document = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            } catch {
                document = XDocument.Parse(xml, LoadOptions.None);
            }

            var root = document.Root;
            if (root == null) {
                return null;
            }

            XNamespace ns = root.Name.Namespace;
            var (workbookPartName, appPartName, corePartName) = ReadRootParts(archive, maxCharactersInPart);
            string? workbookContentType = workbookPartName == null ? null : GetEffectiveContentType(root, workbookPartName);
            if (!IsSupportedWorkbookContentType(workbookContentType)) workbookContentType = WorkbookContentType;
            bool changed = false;

            var xmlDefaults = root.Elements(ns + "Default")
                .Where(e => string.Equals((string?)e.Attribute("Extension"), "xml", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (xmlDefaults.Count != 0) {
                changed = PreserveInheritedXmlContentTypes(archive, root,
                    (string?)xmlDefaults[0].Attribute("ContentType"), xml);
            }

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

            if (workbookPartName != null && EnsureOverride(workbookPartName, workbookContentType!)) changed = true;
            if (appPartName != null && EnsureOverride(appPartName, AppPropsContentType)) changed = true;
            if (corePartName != null && EnsureOverride(corePartName, CorePropsContentType)) changed = true;

            if (!changed) {
                return null;
            }

            document.Declaration = new XDeclaration("1.0", "utf-8", null);

            // Bound the completed declaration before replacing the original ZIP entry.
            using var normalizedTypes = new MemoryStream();
            using var boundedTypes = new ExcelBoundedSeekableStream(
                normalizedTypes, MaximumContentTypesEntryBytes, leaveOpen: true,
                limitExceededExceptionFactory: static _ => new InvalidDataException(
                    "Normalizing package content types exceeds the supported XML size limit."));
            var settings = new XmlWriterSettings {
                Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                Indent = false,
                OmitXmlDeclaration = false,
                NewLineHandling = NewLineHandling.None
            };
            using (var writer = XmlWriter.Create(boundedTypes, settings)) {
                document.Save(writer);
            }
            boundedTypes.Complete();
            return normalizedTypes.ToArray();
        }

        /// <summary>Preserves effective part types before replacing a specialized XML default.</summary>
        private static bool PreserveInheritedXmlContentTypes(
            ZipArchive archive, XElement root, string? previousXmlType, string originalXml) {
            if (string.IsNullOrWhiteSpace(previousXmlType)
                || string.Equals(previousXmlType, "application/xml", StringComparison.OrdinalIgnoreCase)) return false;

            XNamespace ns = root.Name.Namespace;
            var overriddenParts = new HashSet<string>(root.Elements(ns + "Override")
                .Select(item => (string?)item.Attribute("PartName") ?? string.Empty), StringComparer.OrdinalIgnoreCase);
            long remainingBytes = MaximumContentTypesEntryBytes - Encoding.UTF8.GetByteCount(originalXml);
            bool changed = false;
            foreach (ZipArchiveEntry part in archive.Entries) {
                if (part.FullName == ContentTypesEntry
                    || !part.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) continue;
                string partName = "/" + part.FullName;
                if (!overriddenParts.Add(partName)) continue;

                var declaration = new XElement(ns + "Override",
                    new XAttribute("PartName", partName), new XAttribute("ContentType", previousXmlType));
                // Include namespace/escaping overhead conservatively while bounding DOM growth.
                remainingBytes -= Encoding.UTF8.GetByteCount(declaration.ToString(SaveOptions.DisableFormatting));
                if (remainingBytes < 0) throw new InvalidDataException(
                    "Normalizing package content types exceeds the supported XML size limit.");
                root.Add(declaration);
                changed = true;
            }
            return changed;
        }

        internal static bool NeedsContentTypeNormalization(Stream packageStream, long maxCharactersInPart = 0) {
            try {
                using var archive = new ZipArchive(packageStream, ZipArchiveMode.Read, leaveOpen: true);
                var entry = archive.GetEntry(ContentTypesEntry);
                if (entry == null) {
                    return true;
                }

                if (!TryReadContentTypeDeclarations(entry, maxCharactersInPart, out int xmlDefaultCount, out bool xmlDefaultIsCorrect,
                    out List<KeyValuePair<string?, string?>> overrides)) {
                    return true;
                }

                var (workbookPartName, appPartName, corePartName) = ReadRootParts(archive, maxCharactersInPart);
                bool workbookOverrideIsCorrect = workbookPartName == null;
                bool appPropsOverrideIsCorrect = appPartName == null;
                bool corePropsOverrideIsCorrect = corePartName == null;

                foreach (KeyValuePair<string?, string?> declaration in overrides) {
                    string? partName = declaration.Key;
                    string? contentType = declaration.Value;
                    if (string.Equals(partName, workbookPartName, StringComparison.OrdinalIgnoreCase)) {
                        workbookOverrideIsCorrect = IsSupportedWorkbookContentType(contentType);
                    } else if (string.Equals(partName, appPartName, StringComparison.OrdinalIgnoreCase)) {
                        appPropsOverrideIsCorrect = string.Equals(contentType, AppPropsContentType, StringComparison.OrdinalIgnoreCase);
                    } else if (string.Equals(partName, corePartName, StringComparison.OrdinalIgnoreCase)) {
                        corePropsOverrideIsCorrect = string.Equals(contentType, CorePropsContentType, StringComparison.OrdinalIgnoreCase);
                    }
                }

                return xmlDefaultCount != 1
                    || !xmlDefaultIsCorrect
                    || !workbookOverrideIsCorrect
                    || !appPropsOverrideIsCorrect
                    || !corePropsOverrideIsCorrect;
            } catch (InvalidDataException) {
                throw;
            } catch (XmlException exception) {
                throw new InvalidDataException("Package root relationships are invalid or exceed XML limits.", exception);
            } catch {
                return true;
            }
        }

        /// <summary>
        /// Streams the declarations that decide whether a package needs normalizing. Anything this reader rejects,
        /// including a size limit, returns false and leaves the verdict and its error to the normalizer.
        /// </summary>
        private static bool TryReadContentTypeDeclarations(ZipArchiveEntry entry, long maxCharactersInPart,
            out int xmlDefaultCount, out bool xmlDefaultIsCorrect, out List<KeyValuePair<string?, string?>> overrides) {
            xmlDefaultCount = 0;
            xmlDefaultIsCorrect = false;
            overrides = new List<KeyValuePair<string?, string?>>();
            if (entry.Length > MaximumContentTypesEntryBytes) {
                throw new InvalidDataException(
                    $"The package content-types part exceeds the supported {MaximumContentTypesEntryBytes}-byte limit.");
            }

            using Stream input = entry.Open();
            using var reader = XmlReader.Create(input, new XmlReaderSettings {
                DtdProcessing = DtdProcessing.Prohibit,
                XmlResolver = null,
                MaxCharactersInDocument = maxCharactersInPart > 0
                    ? Math.Min(maxCharactersInPart, MaximumContentTypesEntryBytes)
                    : MaximumContentTypesEntryBytes
            });
            try {
                string? rootNamespace = null;
                while (reader.Read()) {
                    if (reader.NodeType != XmlNodeType.Element) continue;
                    if (reader.Depth == 0) { rootNamespace = reader.NamespaceURI; continue; }
                    if (reader.Depth != 1 || reader.NamespaceURI != rootNamespace) continue;
                    if (reader.LocalName == "Default") {
                        if (string.Equals(reader.GetAttribute("Extension"), "xml", StringComparison.OrdinalIgnoreCase)) {
                            xmlDefaultCount++;
                            xmlDefaultIsCorrect = string.Equals(reader.GetAttribute("ContentType"), "application/xml", StringComparison.OrdinalIgnoreCase);
                        }
                    } else if (reader.LocalName == "Override") {
                        overrides.Add(new KeyValuePair<string?, string?>(reader.GetAttribute("PartName"), reader.GetAttribute("ContentType")));
                    }
                }
                return rootNamespace != null;
            } catch (XmlException) {
                return false;
            }
        }

        private static string ReadContentTypesXml(ZipArchiveEntry entry, long maxCharactersInPart) {
            long declaredLength = entry.Length;
            if (declaredLength > MaximumContentTypesEntryBytes) {
                throw new InvalidDataException(
                    $"The package content-types part exceeds the supported {MaximumContentTypesEntryBytes}-byte limit.");
            }

            long characterLimit = maxCharactersInPart > 0
                ? Math.Min(maxCharactersInPart, MaximumContentTypesEntryBytes)
                : MaximumContentTypesEntryBytes;
            using Stream entryStream = entry.Open();
            using var reader = new StreamReader(
                entryStream,
                Encoding.UTF8,
                detectEncodingFromByteOrderMarks: true,
                bufferSize: 4096,
                leaveOpen: false);
            var text = new StringBuilder((int)Math.Min(Math.Min(declaredLength, characterLimit), 8192L));
            var buffer = new char[(int)Math.Min(4096L, characterLimit)];
            while (true) {
                int read = reader.Read(buffer, 0, buffer.Length);
                if (read == 0) return text.ToString();
                if (text.Length > characterLimit - read) {
                    throw new InvalidDataException(
                        $"The package content-types part exceeds the configured {characterLimit}-character limit.");
                }
                text.Append(buffer, 0, read);
            }
        }

        private static bool IsSupportedWorkbookContentType(string? contentType) {
            return string.Equals(contentType, WorkbookContentType, StringComparison.OrdinalIgnoreCase)
                || string.Equals(contentType, MacroEnabledWorkbookContentType, StringComparison.OrdinalIgnoreCase)
                || string.Equals(contentType, TemplateWorkbookContentType, StringComparison.OrdinalIgnoreCase)
                || string.Equals(contentType, MacroEnabledTemplateWorkbookContentType, StringComparison.OrdinalIgnoreCase)
                || string.Equals(contentType, AddInWorkbookContentType, StringComparison.OrdinalIgnoreCase);
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
            string? workbookPartName = GetWorkbookPartName(archive);
            var workbookOverride = root.Elements(ns + "Override")
                .FirstOrDefault(e => string.Equals((string?)e.Attribute("PartName"), workbookPartName, StringComparison.OrdinalIgnoreCase));
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
