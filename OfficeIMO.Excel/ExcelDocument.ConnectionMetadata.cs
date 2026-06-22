using System.Text;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private const string WorkbookConnectionRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
        private const string WorkbookConnectionContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";
        private const string WorksheetQueryTableRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable";
        private const string WorksheetQueryTableContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml";
        private const string WorkbookSlicerCacheRelationshipType = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";
        private const string WorkbookSlicerCacheContentType = "application/vnd.ms-excel.slicerCache+xml";
        private const string WorkbookTimelineCacheRelationshipType = "http://schemas.microsoft.com/office/2011/relationships/timelineCache";
        private const string WorkbookTimelineCacheContentType = "application/vnd.ms-excel.timelineCache+xml";

        /// <summary>
        /// Adds caller-supplied workbook connection metadata XML as a workbook package part.
        /// OfficeIMO preserves this metadata but does not execute or refresh external connections.
        /// </summary>
        /// <param name="xml">Connection metadata XML.</param>
        /// <returns>The added package part.</returns>
        public ExtendedPart AddWorkbookConnectionMetadata(string xml) {
            return AddWorkbookMetadataPart(
                WorkbookConnectionRelationshipType,
                WorkbookConnectionContentType,
                xml);
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
            return AddWorksheetMetadataPart(
                sheet,
                WorksheetQueryTableRelationshipType,
                WorksheetQueryTableContentType,
                xml);
        }

        /// <summary>
        /// Adds workbook-level slicer cache metadata. This authors package metadata for slicer workflows; Excel may still be required to materialize full UI slicer shapes and bindings.
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
        /// Adds workbook-level timeline cache metadata. This authors package metadata for timeline workflows; Excel may still be required to materialize full UI timeline shapes and bindings.
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
            return AddMetadataPart(sheet.WorksheetPart, relationshipType, contentType, xml, targetExtension);
        }

        private ExtendedPart AddMetadataPart(OpenXmlPartContainer container, string relationshipType, string contentType, string xml, string targetExtension) {
            if (container == null) throw new ArgumentNullException(nameof(container));
            if (string.IsNullOrWhiteSpace(relationshipType)) throw new ArgumentNullException(nameof(relationshipType));
            if (string.IsNullOrWhiteSpace(contentType)) throw new ArgumentNullException(nameof(contentType));
            if (string.IsNullOrWhiteSpace(xml)) throw new ArgumentNullException(nameof(xml));

            var part = container.AddExtendedPart(relationshipType, contentType, NormalizeMetadataPartExtension(targetExtension));
            using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write)) {
                byte[] bytes = Encoding.UTF8.GetBytes(xml);
                stream.Write(bytes, 0, bytes.Length);
            }

            _packageDirty = true;
            _packageContentTypesKnownNormalized = true;
            _simplePackageContentKnown = false;
            MarkRequiresSavePreflight();
            return part;
        }

        private static string NormalizeMetadataPartExtension(string targetExtension) {
            if (string.IsNullOrWhiteSpace(targetExtension)) {
                return "xml";
            }

            return targetExtension.Trim().TrimStart('.');
        }
    }
}
