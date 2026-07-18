using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Threading;
using System.Xml;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {

        private sealed class ExtendedWorkbookPackageModel {
            private ExtendedWorkbookPackageModel(
                string workbookRelationshipId,
                IReadOnlyList<ExtendedPartModel> parts,
                IReadOnlyDictionary<OpenXmlPart, ExtendedPartModel> partMap,
                IReadOnlyList<ExtendedRelationshipModel> packageRelationships,
                string corePropertiesXml,
                string appPropertiesXml,
                DirectDataSetWorkbookModel? directDataSetModel,
                IReadOnlyDictionary<OpenXmlPart, DirectDataSetSheetModel> directWorksheetModels) {
                WorkbookRelationshipId = workbookRelationshipId;
                Parts = parts;
                PartMap = partMap;
                PackageRelationships = packageRelationships;
                CorePropertiesXml = corePropertiesXml;
                AppPropertiesXml = appPropertiesXml;
                DirectDataSetModel = directDataSetModel;
                DirectWorksheetModels = directWorksheetModels;
            }

            internal string WorkbookRelationshipId { get; }

            internal IReadOnlyList<ExtendedPartModel> Parts { get; }

            internal IReadOnlyDictionary<OpenXmlPart, ExtendedPartModel> PartMap { get; }

            internal IReadOnlyList<ExtendedRelationshipModel> PackageRelationships { get; }

            internal string CorePropertiesXml { get; }

            internal string AppPropertiesXml { get; }

            internal DirectDataSetWorkbookModel? DirectDataSetModel { get; }

            internal IReadOnlyDictionary<OpenXmlPart, DirectDataSetSheetModel> DirectWorksheetModels { get; }

            internal static bool TryCreate(SpreadsheetDocument document, DirectDataSetWorkbookModel? directDataSetModel, out ExtendedWorkbookPackageModel model, out string? skipReason) {
                model = null!;
                skipReason = null;

                var workbookPart = document.WorkbookPart;
                if (workbookPart?.Workbook == null) {
                    skipReason = "Workbook is missing workbook XML.";
                    return false;
                }

                if (document.HyperlinkRelationships.Any() || document.DataPartReferenceRelationships.Any()) {
                    skipReason = "Package contains reference relationships outside the extended package writer surface.";
                    return false;
                }

                string workbookRelationshipId = document.GetIdOfPart(workbookPart);
                var parts = new List<ExtendedPartModel>();
                var partMap = new Dictionary<OpenXmlPart, ExtendedPartModel>();
                if (!TryCollectPart(workbookPart, parts, partMap, out skipReason)) {
                    return false;
                }

                var packageRelationships = new List<ExtendedRelationshipModel> {
                    new ExtendedRelationshipModel(
                        workbookRelationshipId,
                        workbookPart.RelationshipType,
                        NormalizePackagePartPath(workbookPart.Uri),
                        isExternal: false)
                };

                foreach (var child in document.Parts) {
                    var rootPart = child.OpenXmlPart;
                    if (ReferenceEquals(rootPart, workbookPart) || rootPart is ExtendedFilePropertiesPart) {
                        continue;
                    }

                    if (!TryCollectPart(rootPart, parts, partMap, out skipReason)) {
                        return false;
                    }

                    packageRelationships.Add(new ExtendedRelationshipModel(
                        child.RelationshipId,
                        rootPart.RelationshipType,
                        NormalizePackagePartPath(rootPart.Uri),
                        isExternal: false));
                }

                foreach (var external in document.ExternalRelationships) {
                    packageRelationships.Add(new ExtendedRelationshipModel(
                        external.Id,
                        external.RelationshipType,
                        external.Uri.ToString(),
                        isExternal: true));
                }

                if (!packageRelationships.Any(static relationship => string.Equals(
                    relationship.RelationshipType,
                    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                    StringComparison.Ordinal))) {
                    packageRelationships.Add(new ExtendedRelationshipModel(
                        CreateRelationshipId(packageRelationships),
                        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                        "docProps/core.xml",
                        isExternal: false));
                }

                if (!packageRelationships.Any(static relationship => string.Equals(
                    relationship.RelationshipType,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
                    StringComparison.Ordinal))) {
                    packageRelationships.Add(new ExtendedRelationshipModel(
                        CreateRelationshipId(packageRelationships),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
                        "docProps/app.xml",
                        isExternal: false));
                }

                var directWorksheetModels = directDataSetModel == null
                    ? new Dictionary<OpenXmlPart, DirectDataSetSheetModel>(0)
                    : BuildDirectWorksheetModelMap(workbookPart, directDataSetModel);

                string corePropertiesXml = CreateCorePropertiesXml(document);
                string appPropertiesXml = document.ExtendedFilePropertiesPart?.Properties != null
                    ? "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + document.ExtendedFilePropertiesPart.Properties.OuterXml
                    : CreateAppPropertiesXml();

                model = new ExtendedWorkbookPackageModel(
                    workbookRelationshipId,
                    parts,
                    partMap,
                    packageRelationships,
                    corePropertiesXml,
                    appPropertiesXml,
                    directDataSetModel,
                    directWorksheetModels);
                return true;
            }

            private static IReadOnlyDictionary<OpenXmlPart, DirectDataSetSheetModel> BuildDirectWorksheetModelMap(WorkbookPart workbookPart, DirectDataSetWorkbookModel directDataSetModel) {
                var workbook = workbookPart.Workbook;
                if (directDataSetModel.Sheets.Count == 0 || workbook?.Sheets == null) {
                    return new Dictionary<OpenXmlPart, DirectDataSetSheetModel>(0);
                }

                var directSheetsByName = new Dictionary<string, DirectDataSetSheetModel>(StringComparer.Ordinal);
                for (int i = 0; i < directDataSetModel.Sheets.Count; i++) {
                    DirectDataSetSheetModel sheetModel = directDataSetModel.Sheets[i];
                    directSheetsByName[sheetModel.SheetName] = sheetModel;
                }

                var map = new Dictionary<OpenXmlPart, DirectDataSetSheetModel>();
                foreach (Sheet sheet in workbook.Sheets.Elements<Sheet>()) {
                    if (sheet.Name == null || sheet.Id == null) {
                        continue;
                    }

                    if (!directSheetsByName.TryGetValue(sheet.Name.Value ?? string.Empty, out DirectDataSetSheetModel? sheetModel)) {
                        continue;
                    }

                    if (workbookPart.GetPartById(sheet.Id!) is WorksheetPart worksheetPart) {
                        map[worksheetPart] = sheetModel;
                    }
                }

                return map;
            }

            private static string CreateRelationshipId(IReadOnlyList<ExtendedRelationshipModel> relationships) {
                for (int i = 1; ; i++) {
                    string id = "rId" + InvariantNumberText.Get(i);
                    if (!relationships.Any(relationship => string.Equals(relationship.Id, id, StringComparison.Ordinal))) {
                        return id;
                    }
                }
            }

            private static bool TryCollectPart(
                OpenXmlPart part,
                List<ExtendedPartModel> parts,
                Dictionary<OpenXmlPart, ExtendedPartModel> partMap,
                out string? skipReason) {
                skipReason = null;

                if (partMap.ContainsKey(part)) {
                    return true;
                }

                if (part.DataPartReferenceRelationships.Any()) {
                    skipReason = "Part '" + part.Uri + "' contains data-part relationships outside the extended package writer surface.";
                    return false;
                }

                string path = NormalizePackagePartPath(part.Uri);
                ExtendedPartModel model;
                if (TryCopyRawSupportedPart(part, path, out var rawModel)) {
                    model = rawModel;
                } else {
                    if (!TryGetSupportedRootElement(part, out var rootElement, out skipReason)) {
                        return false;
                    }

                    model = new ExtendedPartModel(part, path, part.ContentType, rootElement, copyRawPart: false);
                }

                parts.Add(model);
                partMap[part] = model;

                foreach (var child in part.Parts) {
                    if (!TryCollectPart(child.OpenXmlPart, parts, partMap, out skipReason)) {
                        return false;
                    }
                }

                return true;
            }
        }

        private sealed class ExtendedPartModel {
            internal ExtendedPartModel(OpenXmlPart part, string path, string contentType, OpenXmlElement? rootElement, bool copyRawPart) {
                Part = part;
                Path = path;
                ContentType = contentType;
                RootElement = rootElement;
                CopyRawPart = copyRawPart;
            }

            internal OpenXmlPart Part { get; }

            internal string Path { get; }

            internal string ContentType { get; }

            internal OpenXmlElement? RootElement { get; }

            internal bool CopyRawPart { get; }
        }

        private sealed class ExtendedRelationshipModel {
            internal ExtendedRelationshipModel(string id, string relationshipType, string target, bool isExternal) {
                Id = id;
                RelationshipType = relationshipType;
                Target = target;
                IsExternal = isExternal;
            }

            internal string Id { get; }

            internal string RelationshipType { get; }

            internal string Target { get; }

            internal bool IsExternal { get; }
        }

        private static bool TryGetSupportedRootElement(OpenXmlPart part, out OpenXmlElement rootElement, out string? skipReason) {
            skipReason = null;
            rootElement = part switch {
                WorkbookPart workbookPart when workbookPart.Workbook != null => workbookPart.Workbook,
                WorksheetPart worksheetPart when worksheetPart.Worksheet != null => worksheetPart.Worksheet,
                WorkbookStylesPart stylesPart when stylesPart.Stylesheet != null => stylesPart.Stylesheet,
                SharedStringTablePart sharedStringPart when sharedStringPart.SharedStringTable != null => sharedStringPart.SharedStringTable,
                CustomFilePropertiesPart customPropertiesPart when customPropertiesPart.Properties != null => customPropertiesPart.Properties,
                ThemePart themePart when themePart.Theme != null => themePart.Theme,
                DrawingsPart drawingsPart when drawingsPart.WorksheetDrawing != null => drawingsPart.WorksheetDrawing,
                ChartPart chartPart when chartPart.ChartSpace != null => chartPart.ChartSpace,
                TableDefinitionPart tablePart when tablePart.Table != null => tablePart.Table,
                PivotTablePart pivotTablePart when pivotTablePart.PivotTableDefinition != null => pivotTablePart.PivotTableDefinition,
                PivotTableCacheDefinitionPart cacheDefinitionPart when cacheDefinitionPart.PivotCacheDefinition != null => cacheDefinitionPart.PivotCacheDefinition,
                PivotTableCacheRecordsPart cacheRecordsPart when cacheRecordsPart.PivotCacheRecords != null => cacheRecordsPart.PivotCacheRecords,
                WorksheetThreadedCommentsPart threadedCommentsPart when threadedCommentsPart.ThreadedComments != null => threadedCommentsPart.ThreadedComments,
                WorkbookPersonPart personPart when personPart.PersonList != null => personPart.PersonList,
                _ => null!
            };

            if (rootElement == null) {
                skipReason = "Part '" + part.Uri + "' is outside the extended package writer surface.";
                return false;
            }

            if (rootElement.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                skipReason = "Part '" + part.Uri + "' contains unknown Open XML elements.";
                return false;
            }

            return true;
        }

        private sealed class RawPivotCacheRecordsPartMarker {
        }

        private static readonly System.Runtime.CompilerServices.ConditionalWeakTable<PivotTableCacheRecordsPart, RawPivotCacheRecordsPartMarker> RawPivotCacheRecordsParts = new();

        internal static void MarkPivotCacheRecordsPartAsRawWritten(PivotTableCacheRecordsPart part)
            => RawPivotCacheRecordsParts.GetValue(part, static _ => new RawPivotCacheRecordsPartMarker());

        internal static void MarkPivotCacheRecordsPartAsModelWritten(PivotTableCacheRecordsPart part)
            => RawPivotCacheRecordsParts.Remove(part);

        private static bool IsRawPivotCacheRecordsPart(OpenXmlPart part)
            => part is PivotTableCacheRecordsPart cacheRecordsPart
               && RawPivotCacheRecordsParts.TryGetValue(cacheRecordsPart, out _);

        private static bool TryCopyRawSupportedPart(OpenXmlPart part, string path, out ExtendedPartModel model) {
            model = null!;
            if (part.IsRootElementLoaded
                && part is not ChartStylePart
                && part is not ChartColorStylePart
                && !IsRawPivotCacheRecordsPart(part)) {
                return false;
            }

            using (var source = part.GetStream(FileMode.Open, FileAccess.Read)) {
                if (source.CanSeek && source.Length == 0) {
                    return false;
                }
            }

            model = new ExtendedPartModel(part, path, part.ContentType, rootElement: null, copyRawPart: true);
            return true;
        }

        private static string NormalizePackagePartPath(Uri uri) {
            string path = uri.OriginalString.Replace('\\', '/');
            if (path.StartsWith("/", StringComparison.Ordinal)) {
                path = path.Substring(1);
            }

            return path;
        }

        private static string GetRelationshipsPath(string partPath) {
            int slash = partPath.LastIndexOf('/');
            if (slash < 0) {
                return "_rels/" + partPath + ".rels";
            }

            return partPath.Substring(0, slash + 1) + "_rels/" + partPath.Substring(slash + 1) + ".rels";
        }

        private static string GetRelativeTargetPath(string sourcePath, string targetPath) {
            int slash = sourcePath.LastIndexOf('/');
            string sourceDirectory = slash < 0 ? string.Empty : sourcePath.Substring(0, slash + 1);
            var sourceUri = new Uri("x:///" + sourceDirectory, UriKind.Absolute);
            var targetUri = new Uri("x:///" + targetPath, UriKind.Absolute);
            return Uri.UnescapeDataString(sourceUri.MakeRelativeUri(targetUri).ToString());
        }

        private static void WriteExtendedContentTypesEntry(ZipArchive archive, IReadOnlyList<ExtendedPartModel> parts, bool includeDirectStyles, bool includeDirectSharedStrings) {
            var builder = new System.Text.StringBuilder(768 + parts.Count * 180);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            builder.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            builder.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            builder.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
            builder.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
            if (includeDirectStyles && !parts.Any(static part => string.Equals(part.Path, "xl/styles.xml", StringComparison.Ordinal))) {
                builder.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            }

            if (includeDirectSharedStrings && !parts.Any(static part => string.Equals(part.Path, "xl/sharedStrings.xml", StringComparison.Ordinal))) {
                builder.Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
            }

            foreach (var part in parts.OrderBy(static item => item.Path, StringComparer.Ordinal)) {
                builder.Append("<Override PartName=\"/");
                AppendXmlEscaped(builder, part.Path);
                builder.Append("\" ContentType=\"");
                AppendXmlEscaped(builder, part.ContentType);
                builder.Append("\"/>");
            }

            builder.Append("</Types>");
            WriteTextEntry(archive, "[Content_Types].xml", builder.ToString());
        }

        private static void WriteExtendedRelationshipsEntry(ZipArchive archive, string path, IReadOnlyList<ExtendedRelationshipModel> relationships) {
            var builder = new System.Text.StringBuilder(160 + relationships.Count * 220);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var relationship in relationships.OrderBy(static item => item.Id, StringComparer.Ordinal)) {
                builder.Append("<Relationship Id=\"");
                AppendXmlEscaped(builder, relationship.Id);
                builder.Append("\" Type=\"");
                AppendXmlEscaped(builder, relationship.RelationshipType);
                builder.Append("\" Target=\"");
                AppendXmlEscaped(builder, relationship.Target);
                builder.Append('"');
                if (relationship.IsExternal) {
                    builder.Append(" TargetMode=\"External\"");
                }

                builder.Append("/>");
            }

            builder.Append("</Relationships>");
            WriteTextEntry(archive, path, builder.ToString());
        }
    }
}
