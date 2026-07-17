using System.Globalization;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private const string WorkbookConnectionRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
        private const string WorkbookConnectionContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
        private const string WorksheetQueryTableRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable";
        private const string WorksheetQueryTableContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
        private const string SpreadsheetMainNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string OfficeDocumentRelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string OfficeImoPivotInteractionNamespace = "https://schemas.evotec.xyz/officeimo/excel";
        private const string MicrosoftWorkbookSlicerCacheRelationshipType = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
        private const string MicrosoftWorkbookSlicerCacheContentType = "application/vnd.ms-excel.slicerCache+xml";
        private const string MicrosoftWorkbookTimelineCacheRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/timelineCache";
        private const string MicrosoftWorkbookTimelineCacheContentType = "application/vnd.ms-excel.timelineCache+xml";
        private const string WorkbookSlicerCacheRelationshipType = "https://schemas.evotec.xyz/officeimo/excel/relationships/slicerCacheMetadata";
        private const string WorkbookSlicerCacheContentType = "application/vnd.officeimo.excel.slicerCache-metadata+xml";
        private const string WorkbookTimelineCacheRelationshipType = "https://schemas.evotec.xyz/officeimo/excel/relationships/timelineCacheMetadata";
        private const string WorkbookTimelineCacheContentType = "application/vnd.officeimo.excel.timelineCache-metadata+xml";

        /// <summary>
        /// Adds caller-supplied workbook connection metadata XML as a workbook package part.
        /// OfficeIMO preserves this metadata but does not execute or refresh external connections.
        /// </summary>
        /// <param name="xml">Connection metadata XML.</param>
        /// <returns>The added or updated package part.</returns>
        public OpenXmlPart AddWorkbookConnectionMetadata(string xml) {
            OpenXmlPart? existingPart = GetWorkbookConnectionPart();
            if (existingPart == null) {
                return AddWorkbookMetadataPart(
                    WorkbookConnectionRelationshipType,
                    WorkbookConnectionContentType,
                    NormalizeWorkbookConnectionMetadata(xml));
            }

            string mergedXml = MergeWorkbookConnectionMetadata(ReadMetadataPart(existingPart), xml);
            WriteMetadataPart(existingPart, mergedXml);
            MarkMetadataPartChanged();
            return existingPart;
        }

        /// <summary>
        /// Adds caller-supplied query-table metadata XML as a worksheet package part.
        /// OfficeIMO preserves this metadata but does not execute or refresh external queries.
        /// </summary>
        /// <param name="worksheetName">Worksheet that owns the query-table metadata.</param>
        /// <param name="xml">Query-table metadata XML.</param>
        /// <returns>The added package part.</returns>
        public ExtendedPart AddWorksheetQueryTableMetadata(string worksheetName, string xml) {
            if (string.IsNullOrWhiteSpace(worksheetName)) throw new ArgumentNullException(nameof(worksheetName));
            var sheet = this[worksheetName];
            ExtendedPart part = AddWorksheetMetadataPart(
                sheet,
                WorksheetQueryTableRelationshipType,
                WorksheetQueryTableContentType,
                xml);
            LinkWorksheetQueryTablePart(sheet.WorksheetPart, part);
            return part;
        }

        /// <summary>
        /// Adds OfficeIMO-owned workbook-level slicer binding metadata. This does not impersonate a native Excel slicer-cache part;
        /// native cache structures and UI shapes must be materialized separately.
        /// </summary>
        /// <param name="options">Slicer cache metadata options.</param>
        /// <returns>The added package part.</returns>
        public ExtendedPart AddWorkbookSlicerCache(ExcelSlicerCacheOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            return AddWorkbookMetadataPart(
                WorkbookSlicerCacheRelationshipType,
                WorkbookSlicerCacheContentType,
                options.ToXml());
        }

        /// <summary>
        /// Adds OfficeIMO-owned workbook-level timeline binding metadata. This does not impersonate a native Excel timeline-cache part;
        /// native cache structures and UI shapes must be materialized separately.
        /// </summary>
        /// <param name="options">Timeline cache metadata options.</param>
        /// <returns>The added package part.</returns>
        public ExtendedPart AddWorkbookTimelineCache(ExcelTimelineCacheOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            return AddWorkbookMetadataPart(
                WorkbookTimelineCacheRelationshipType,
                WorkbookTimelineCacheContentType,
                options.ToXml());
        }

        /// <summary>
        /// Adds caller-supplied XML as a workbook package metadata part.
        /// </summary>
        /// <param name="relationshipType">Open XML relationship type.</param>
        /// <param name="contentType">Package part content type.</param>
        /// <param name="xml">Metadata XML.</param>
        /// <param name="targetExtension">Target extension for the generated package part.</param>
        /// <returns>The added package part.</returns>
        public ExtendedPart AddWorkbookMetadataPart(string relationshipType, string contentType, string xml, string targetExtension = "xml") {
            return AddMetadataPart(WorkbookPartRoot, relationshipType, contentType, xml, targetExtension);
        }

        /// <summary>
        /// Adds caller-supplied XML as a worksheet package metadata part.
        /// </summary>
        /// <param name="sheet">Worksheet that owns the metadata part.</param>
        /// <param name="relationshipType">Open XML relationship type.</param>
        /// <param name="contentType">Package part content type.</param>
        /// <param name="xml">Metadata XML.</param>
        /// <param name="targetExtension">Target extension for the generated package part.</param>
        /// <returns>The added package part.</returns>
        public ExtendedPart AddWorksheetMetadataPart(ExcelSheet sheet, string relationshipType, string contentType, string xml, string targetExtension = "xml") {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            if (!ReferenceEquals(sheet.Document, this)) {
                throw new ArgumentException("Worksheet metadata can only be added to a worksheet owned by this workbook.", nameof(sheet));
            }

            return AddMetadataPart(sheet.WorksheetPart, relationshipType, contentType, xml, targetExtension);
        }

        private ExtendedPart AddMetadataPart(OpenXmlPartContainer container, string relationshipType, string contentType, string xml, string targetExtension) {
            if (container == null) throw new ArgumentNullException(nameof(container));
            if (string.IsNullOrWhiteSpace(relationshipType)) throw new ArgumentNullException(nameof(relationshipType));
            if (string.IsNullOrWhiteSpace(contentType)) throw new ArgumentNullException(nameof(contentType));
            if (string.IsNullOrWhiteSpace(xml)) throw new ArgumentNullException(nameof(xml));

            var part = container.AddExtendedPart(relationshipType, contentType, NormalizeMetadataPartExtension(targetExtension));
            WriteMetadataPart(part, xml);
            MarkMetadataPartChanged();
            return part;
        }

        private void LinkWorksheetQueryTablePart(WorksheetPart worksheetPart, OpenXmlPart queryTablePart) {
            string relationshipId = worksheetPart.GetIdOfPart(queryTablePart);
            Worksheet worksheet = worksheetPart.Worksheet ?? throw new InvalidDataException("Worksheet metadata cannot be linked because the worksheet part has no worksheet XML.");
            OpenXmlElement? queryTableParts = worksheet.ChildElements.FirstOrDefault(IsQueryTablePartsElement);
            if (queryTableParts == null) {
                queryTableParts = new OpenXmlUnknownElement(string.Empty, "queryTableParts", SpreadsheetMainNamespace);
                WorksheetExtensionList? extensionList = worksheet.GetFirstChild<WorksheetExtensionList>();
                if (extensionList != null) {
                    worksheet.InsertBefore(queryTableParts, extensionList);
                } else {
                    worksheet.Append(queryTableParts);
                }
            }

            bool exists = queryTableParts.ChildElements
                .Where(IsQueryTablePartElement)
                .Any(part => string.Equals(GetRelationshipId(part), relationshipId, StringComparison.Ordinal));
            if (!exists) {
                var linkedPart = new OpenXmlUnknownElement(string.Empty, "queryTablePart", SpreadsheetMainNamespace);
                linkedPart.SetAttribute(new OpenXmlAttribute("r", "id", OfficeDocumentRelationshipsNamespace, relationshipId));
                queryTableParts.Append(linkedPart);
            }

            int count = queryTableParts.ChildElements.Count(IsQueryTablePartElement);
            queryTableParts.SetAttribute(new OpenXmlAttribute("count", string.Empty, count.ToString(CultureInfo.InvariantCulture)));
            worksheet.Save();
            MarkMetadataPartChanged();
        }

        private static bool IsQueryTablePartsElement(OpenXmlElement element) {
            return string.Equals(element.LocalName, "queryTableParts", StringComparison.Ordinal)
                && string.Equals(element.NamespaceUri, SpreadsheetMainNamespace, StringComparison.Ordinal);
        }

        private static bool IsQueryTablePartElement(OpenXmlElement element) {
            return string.Equals(element.LocalName, "queryTablePart", StringComparison.Ordinal)
                && string.Equals(element.NamespaceUri, SpreadsheetMainNamespace, StringComparison.Ordinal);
        }

        private static string? GetRelationshipId(OpenXmlElement element) {
            string? id = element.GetAttribute("id", OfficeDocumentRelationshipsNamespace).Value;
            if (!string.IsNullOrWhiteSpace(id)) {
                return id;
            }

            return element.GetAttribute("id", string.Empty).Value;
        }

        private OpenXmlPart? GetWorkbookConnectionPart() {
            return EnumerateWorkbookConnectionParts()
                .FirstOrDefault(part => string.Equals(part.ContentType, WorkbookConnectionContentType, StringComparison.OrdinalIgnoreCase)
                    || part is ConnectionsPart);
        }

        private IEnumerable<OpenXmlPart> EnumerateWorkbookConnectionParts() {
            foreach (IdPartPair pair in WorkbookPartRoot.Parts) {
                OpenXmlPart part = pair.OpenXmlPart;
                if (part is ConnectionsPart
                    || string.Equals(part.RelationshipType, WorkbookConnectionRelationshipType, StringComparison.Ordinal)
                    || part.ContentType.IndexOf("connections", StringComparison.OrdinalIgnoreCase) >= 0) {
                    yield return part;
                }
            }
        }

        private static string MergeWorkbookConnectionMetadata(string existingXml, string newXml) {
            XDocument existingDocument = XDocument.Parse(existingXml);
            XDocument newDocument = XDocument.Parse(newXml);
            if (existingDocument.Root == null || newDocument.Root == null) {
                throw new InvalidDataException("Workbook connection metadata must have a document root.");
            }

            IEnumerable<XElement> newConnections = newDocument.Root.Name.LocalName == "connection"
                ? new[] { newDocument.Root }
                : newDocument.Root.Elements().Where(element => element.Name.LocalName == "connection");

            foreach (XElement connection in newConnections) {
                existingDocument.Root.Add(new XElement(connection));
            }

            existingDocument.Root.SetAttributeValue("count", existingDocument.Root.Elements().Count(element => element.Name.LocalName == "connection").ToString(System.Globalization.CultureInfo.InvariantCulture));
            return existingDocument.ToString(SaveOptions.DisableFormatting);
        }

        private static string NormalizeWorkbookConnectionMetadata(string xml) {
            XDocument document = XDocument.Parse(xml);
            if (document.Root == null) {
                throw new InvalidDataException("Workbook connection metadata must have a document root.");
            }

            if (document.Root.Name.LocalName != "connection") {
                return document.ToString(SaveOptions.DisableFormatting);
            }

            XNamespace ns = document.Root.Name.Namespace;
            var connections = new XElement(ns + "connections", new XElement(document.Root));
            connections.SetAttributeValue("count", "1");
            return new XDocument(connections).ToString(SaveOptions.DisableFormatting);
        }

        private static string ReadMetadataPart(OpenXmlPart part) {
            return ReadOpenXmlPartText(part);
        }

        private static void WriteMetadataPart(OpenXmlPart part, string xml) {
            WriteOpenXmlPartText(part, xml);
        }

        private static string ReadOpenXmlPartText(OpenXmlPart part) {
            if (part is ConnectionsPart connectionsPart && connectionsPart.Connections != null) {
                return connectionsPart.Connections.OuterXml;
            }

            using Stream stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var reader = new StreamReader(stream, Encoding.UTF8);
            return reader.ReadToEnd();
        }

        private static void WriteOpenXmlPartText(OpenXmlPart part, string xml) {
            using Stream stream = part.GetStream(FileMode.Create, FileAccess.Write);
            byte[] bytes = Encoding.UTF8.GetBytes(xml);
            stream.Write(bytes, 0, bytes.Length);
            if (part is ConnectionsPart connectionsPart) {
                connectionsPart.Connections = new Connections(xml);
            }
        }

        private void MarkMetadataPartChanged() {
            _packageDirty = true;
            _packageContentTypesKnownNormalized = true;
            _simplePackageContentKnown = false;
            MarkRequiresSavePreflight();
        }

        private static string NormalizeMetadataPartExtension(string targetExtension) {
            if (string.IsNullOrWhiteSpace(targetExtension)) {
                return "xml";
            }

            return targetExtension.Trim().TrimStart('.');
        }
    }
}
