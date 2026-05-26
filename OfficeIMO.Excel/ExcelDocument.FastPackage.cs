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
        private bool TryWriteSimpleWorkbookPackage(Stream destination, ExcelSaveOptions? options, bool updateDocumentState, out string? skipReason, CancellationToken ct = default) {
            skipReason = null;

            if (destination == null || !destination.CanWrite || !destination.CanSeek) {
                skipReason = "Destination stream must be writable and seekable.";
                return false;
            }

            ct.ThrowIfCancellationRequested();

            if (options?.DisableFastPackageWriter == true) {
                skipReason = "Fast package writer was disabled by save options.";
                return false;
            }

            if (options?.ValidateOpenXml == true) {
                skipReason = "Open XML validation requires the standard package finalization path.";
                return false;
            }

            if (_packagePropertiesDirty) {
                skipReason = "Package properties changed.";
                return false;
            }

            if (_unchangedPackageBytes != null) {
                skipReason = "An unchanged package payload is already available.";
                return false;
            }

            if (_packageContentTypesKnownNormalized && !_simplePackageContentKnown) {
                skipReason = "Workbook was loaded or previously finalized; standard save preserves package metadata and relationships.";
                return false;
            }

            if (HasCalculationSaveWork(options)) {
                skipReason = "Calculation save work is pending.";
                return false;
            }

            if (!FastWorkbookPackageModel.TryCreate(_spreadSheetDocument, out var model, out string? modelSkipReason)) {
                skipReason = modelSkipReason ?? "Workbook contains parts or worksheet features outside the simple package writer surface.";
                return false;
            }

            ct.ThrowIfCancellationRequested();
            PrepareDestinationStreamForWrite(destination);
            FastWorkbookPackageWriter.Write(destination, model, ct);

            destination.Flush();
            destination.Seek(0, SeekOrigin.Begin);
            if (updateDocumentState) {
                _packageDirty = false;
                _packagePropertiesDirty = false;
                _requiresSavePreflight = false;
                _unchangedPackageBytes = null;
                _packageContentTypesKnownNormalized = true;
                _simplePackageContentKnown = true;
            }

            return true;
        }

        private bool TryWriteExtendedWorkbookPackage(Stream destination, ExcelSaveOptions? options, bool updateDocumentState, out string? skipReason, CancellationToken ct = default) {
            skipReason = null;

            if (destination == null || !destination.CanWrite || !destination.CanSeek) {
                skipReason = "Destination stream must be writable and seekable.";
                return false;
            }

            ct.ThrowIfCancellationRequested();

            if (options?.DisableFastPackageWriter == true) {
                skipReason = "Fast package writer was disabled by save options.";
                return false;
            }

            if (options?.ValidateOpenXml == true) {
                skipReason = "Open XML validation requires the standard package finalization path.";
                return false;
            }

            if (_packagePropertiesDirty) {
                skipReason = "Package properties changed.";
                return false;
            }

            if (_unchangedPackageBytes != null) {
                skipReason = "An unchanged package payload is already available.";
                return false;
            }

            if (_packageContentTypesKnownNormalized && !_simplePackageContentKnown) {
                skipReason = "Workbook was loaded or previously finalized; standard save preserves package metadata and relationships.";
                return false;
            }

            if (HasCalculationSaveWork(options)) {
                skipReason = "Calculation save work is pending.";
                return false;
            }

            Stopwatch? stageWatch = Execution.OnTiming == null ? null : Stopwatch.StartNew();
            if (!TryRefreshMaterializedDirectDataSetFastSaveModel(out string? directModelSkipReason)) {
                skipReason = directModelSkipReason ?? "Direct worksheet metadata could not be refreshed.";
                return false;
            }

            if (!ExtendedWorkbookPackageModel.TryCreate(_spreadSheetDocument, _materializedDirectDataSetFastSaveModel, out var model, out string? modelSkipReason)) {
                skipReason = modelSkipReason ?? "Workbook contains parts outside the extended package writer surface.";
                return false;
            }
            ReportExtendedPackageTiming(stageWatch, "Save.ExtendedPackage.CreateModel");

            if (ShouldMaterializeMixedDirectWorkbookGlobalParts(model)
                && _materializedDirectDataSetFastSaveModel != null) {
                _materializingDeferredDataSetImport = true;
                try {
                    MaterializeDirectDataSetModel(_materializedDirectDataSetFastSaveModel);
                } finally {
                    _materializingDeferredDataSetImport = false;
                }

                _materializedDirectDataSetFastSaveModel = null;
                if (!ExtendedWorkbookPackageModel.TryCreate(_spreadSheetDocument, directDataSetModel: null, out model, out modelSkipReason)) {
                    skipReason = modelSkipReason ?? "Workbook contains parts outside the extended package writer surface after materializing direct data.";
                    return false;
                }

                ReportExtendedPackageTiming(stageWatch, "Save.ExtendedPackage.MaterializeMixedWorkbookDirectData");
            }

            ct.ThrowIfCancellationRequested();
            PrepareDestinationStreamForWrite(destination);
            ReportExtendedPackageTiming(stageWatch, "Save.ExtendedPackage.PrepareDestination");
            ExtendedWorkbookPackageWriter.Write(destination, model, ct, Execution);
            ReportExtendedPackageTiming(stageWatch, "Save.ExtendedPackage.WritePackage");

            destination.Flush();
            destination.Seek(0, SeekOrigin.Begin);
            ReportExtendedPackageTiming(stageWatch, "Save.ExtendedPackage.FlushAndSeek");
            if (updateDocumentState) {
                _packageDirty = false;
                _packagePropertiesDirty = false;
                _requiresSavePreflight = false;
                _unchangedPackageBytes = null;
                _packageContentTypesKnownNormalized = true;
                _simplePackageContentKnown = true;
            }

            return true;
        }

        private void ReportExtendedPackageTiming(Stopwatch? stopwatch, string operation) {
            if (stopwatch == null) {
                return;
            }

            Execution.ReportTiming(operation, stopwatch.Elapsed);
            stopwatch.Restart();
        }

        private static bool ShouldMaterializeMixedDirectWorkbookGlobalParts(ExtendedWorkbookPackageModel model) {
            if (model.DirectDataSetModel == null || model.DirectWorksheetModels.Count == 0) {
                return false;
            }

            int worksheetPartCount = 0;
            bool hasWorkbookGlobalPart = false;
            foreach (var part in model.Parts) {
                if (part.Part is WorksheetPart) {
                    worksheetPartCount++;
                } else if (part.Part is WorkbookStylesPart || part.Part is SharedStringTablePart) {
                    hasWorkbookGlobalPart = true;
                }
            }

            return hasWorkbookGlobalPart
                   && model.DirectWorksheetModels.Count < worksheetPartCount;
        }

        private static readonly System.Text.UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

        private sealed class ExtendedWorkbookPackageModel {
            private ExtendedWorkbookPackageModel(
                string workbookRelationshipId,
                IReadOnlyList<ExtendedPartModel> parts,
                IReadOnlyDictionary<OpenXmlPart, ExtendedPartModel> partMap,
                IReadOnlyList<ExtendedRelationshipModel> packageRelationships,
                DirectDataSetWorkbookModel? directDataSetModel,
                IReadOnlyDictionary<OpenXmlPart, DirectDataSetSheetModel> directWorksheetModels) {
                WorkbookRelationshipId = workbookRelationshipId;
                Parts = parts;
                PartMap = partMap;
                PackageRelationships = packageRelationships;
                DirectDataSetModel = directDataSetModel;
                DirectWorksheetModels = directWorksheetModels;
            }

            internal string WorkbookRelationshipId { get; }

            internal IReadOnlyList<ExtendedPartModel> Parts { get; }

            internal IReadOnlyDictionary<OpenXmlPart, ExtendedPartModel> PartMap { get; }

            internal IReadOnlyList<ExtendedRelationshipModel> PackageRelationships { get; }

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
                        isExternal: false),
                    new ExtendedRelationshipModel(
                        "rIdCore",
                        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
                        "docProps/core.xml",
                        isExternal: false),
                    new ExtendedRelationshipModel(
                        "rIdApp",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
                        "docProps/app.xml",
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

                var directWorksheetModels = directDataSetModel == null
                    ? new Dictionary<OpenXmlPart, DirectDataSetSheetModel>(0)
                    : BuildDirectWorksheetModelMap(workbookPart, directDataSetModel);

                model = new ExtendedWorkbookPackageModel(workbookRelationshipId, parts, partMap, packageRelationships, directDataSetModel, directWorksheetModels);
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

            private static bool TryCollectPart(
                OpenXmlPart part,
                List<ExtendedPartModel> parts,
                Dictionary<OpenXmlPart, ExtendedPartModel> partMap,
                out string? skipReason) {
                skipReason = null;

                if (partMap.ContainsKey(part)) {
                    return true;
                }

                if (!TryGetSupportedRootElement(part, out var rootElement, out skipReason)) {
                    return false;
                }

                string path = NormalizePackagePartPath(part.Uri);
                var model = new ExtendedPartModel(part, path, part.ContentType, rootElement);
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
            internal ExtendedPartModel(OpenXmlPart part, string path, string contentType, OpenXmlElement rootElement) {
                Part = part;
                Path = path;
                ContentType = contentType;
                RootElement = rootElement;
            }

            internal OpenXmlPart Part { get; }

            internal string Path { get; }

            internal string ContentType { get; }

            internal OpenXmlElement RootElement { get; }
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

        private static class ExtendedWorkbookPackageWriter {
            internal static void Write(Stream destination, ExtendedWorkbookPackageModel model, CancellationToken ct, ExecutionPolicy? execution = null) {
                Stopwatch? stageWatch = execution?.OnTiming == null ? null : Stopwatch.StartNew();
                void ReportTiming(string operation) {
                    if (stageWatch == null || execution == null) {
                        return;
                    }

                    execution.ReportTiming(operation, stageWatch.Elapsed);
                    stageWatch.Restart();
                }

                using var archive = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);
                DirectDataSetWorkbookWriter.ExtendedDirectWritePlan? directWritePlan = null;
                if (CanUseDirectWorksheetEntries(model)
                    && DirectDataSetWorkbookWriter.TryCreateExtendedWritePlan(model.DirectDataSetModel!, ct, out var candidateDirectWritePlan)) {
                    directWritePlan = candidateDirectWritePlan;
                }
                ReportTiming("Save.ExtendedPackage.CreateDirectWritePlan");

                WriteExtendedContentTypesEntry(archive, model.Parts, directWritePlan != null, directWritePlan?.HasSharedStrings == true);
                WriteExtendedRelationshipsEntry(archive, "_rels/.rels", model.PackageRelationships);
                WriteCorePropertiesEntry(archive);
                WriteAppPropertiesEntry(archive);
                ReportTiming("Save.ExtendedPackage.WriteFixedEntries");

                bool wroteDirectStyles = false;
                bool wroteDirectSharedStrings = false;
                foreach (var part in model.Parts) {
                    ct.ThrowIfCancellationRequested();
                    if (directWritePlan != null && part.Part is WorkbookStylesPart) {
                        DirectDataSetWorkbookWriter.WriteExtendedStyles(archive, directWritePlan);
                        wroteDirectStyles = true;
                        ReportTiming("Save.ExtendedPackage.WriteDirectStyles");
                    } else if (directWritePlan?.HasSharedStrings == true && part.Part is SharedStringTablePart) {
                        DirectDataSetWorkbookWriter.WriteExtendedSharedStrings(archive, directWritePlan);
                        wroteDirectSharedStrings = true;
                        ReportTiming("Save.ExtendedPackage.WriteDirectSharedStrings");
                    } else if (directWritePlan != null
                        && part.Part is WorksheetPart directWorksheetPart
                        && model.DirectWorksheetModels.TryGetValue(directWorksheetPart, out DirectDataSetSheetModel? directSheetModel)) {
                        string? tableRelationshipId = null;
                        if (directSheetModel.HasTable) {
                            var tablePart = directWorksheetPart.TableDefinitionParts.FirstOrDefault();
                            if (tablePart != null) {
                                tableRelationshipId = directWorksheetPart.GetIdOfPart(tablePart);
                            }
                        }

                        DirectDataSetWorkbookWriter.WriteExtendedWorksheet(archive, directWritePlan, directSheetModel, part.Path, tableRelationshipId, ct);
                        ReportTiming("Save.ExtendedPackage.WriteDirectWorksheet");
                    } else if (part.Part is WorksheetPart worksheetPart
                        && CanWriteSimpleWorksheet(worksheetPart, worksheetPart.Worksheet!, out _, allowDrawings: true, allowPivotTables: true)) {
                        var tablePartIds = worksheetPart.TableDefinitionParts
                            .Select(worksheetPart.GetIdOfPart)
                            .ToDictionary(static id => id, static id => string.Empty, StringComparer.Ordinal);
                        var worksheetModel = new FastWorksheetPackageModel(
                            string.Empty,
                            0U,
                            null,
                            string.Empty,
                            part.Path,
                            GetRelationshipsPath(part.Path),
                            worksheetPart.Worksheet!,
                            tablePartIds,
                            Array.Empty<FastHyperlinkRelationshipModel>());
                        WriteWorksheetEntry(archive, worksheetModel);
                        ReportTiming("Save.ExtendedPackage.WriteSimpleWorksheet");
                    } else {
                        WriteOpenXmlElementEntry(archive, part.Path, part.RootElement);
                        ReportTiming("Save.ExtendedPackage.WriteOpenXmlPart");
                    }

                    var relationships = CreateRelationships(part.Part, part.Path, model.PartMap, directWritePlan);
                    if (relationships.Count != 0) {
                        WriteExtendedRelationshipsEntry(archive, GetRelationshipsPath(part.Path), relationships);
                        ReportTiming("Save.ExtendedPackage.WriteRelationships");
                    }
                }

                if (directWritePlan != null && !wroteDirectStyles) {
                    DirectDataSetWorkbookWriter.WriteExtendedStyles(archive, directWritePlan);
                    ReportTiming("Save.ExtendedPackage.WriteDirectStyles");
                }

                if (directWritePlan?.HasSharedStrings == true && !wroteDirectSharedStrings) {
                    DirectDataSetWorkbookWriter.WriteExtendedSharedStrings(archive, directWritePlan);
                    ReportTiming("Save.ExtendedPackage.WriteDirectSharedStrings");
                }
            }

            private static bool CanUseDirectWorksheetEntries(ExtendedWorkbookPackageModel model) {
                if (model.DirectDataSetModel == null
                    || model.DirectDataSetModel.Sheets.Count == 0
                    || model.DirectWorksheetModels.Count != model.DirectDataSetModel.Sheets.Count) {
                    return false;
                }

                int worksheetPartCount = 0;
                foreach (var part in model.Parts) {
                    if (part.Part is WorksheetPart) {
                        worksheetPartCount++;
                    }
                }

                if (model.DirectWorksheetModels.Count != worksheetPartCount) {
                    return false;
                }

                foreach (var pair in model.DirectWorksheetModels) {
                    if (pair.Key is not WorksheetPart worksheetPart) {
                        return false;
                    }

                    DirectDataSetSheetModel sheetModel = pair.Value;
                    if (!sheetModel.HasTable) {
                        continue;
                    }

                    var tablePart = worksheetPart.TableDefinitionParts.FirstOrDefault();
                    if (tablePart == null) {
                        return false;
                    }
                }

                return true;
            }

            private static IReadOnlyList<ExtendedRelationshipModel> CreateRelationships(
                OpenXmlPartContainer container,
                string sourcePath,
                IReadOnlyDictionary<OpenXmlPart, ExtendedPartModel> partMap,
                DirectDataSetWorkbookWriter.ExtendedDirectWritePlan? directWritePlan) {
                var relationships = new List<ExtendedRelationshipModel>();
                foreach (var child in container.Parts) {
                    if (!partMap.TryGetValue(child.OpenXmlPart, out var targetPart)) {
                        continue;
                    }

                    relationships.Add(new ExtendedRelationshipModel(
                        child.RelationshipId,
                        child.OpenXmlPart.RelationshipType,
                        GetRelativeTargetPath(sourcePath, targetPart.Path),
                        isExternal: false));
                }

                if (directWritePlan != null
                    && container is WorkbookPart
                    && !relationships.Any(static relationship => string.Equals(
                        relationship.RelationshipType,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                        StringComparison.Ordinal))) {
                    relationships.Add(new ExtendedRelationshipModel(
                        CreateRelationshipId(relationships),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
                        "styles.xml",
                        isExternal: false));
                }

                if (directWritePlan?.HasSharedStrings == true
                    && container is WorkbookPart
                    && !relationships.Any(static relationship => string.Equals(
                        relationship.RelationshipType,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                        StringComparison.Ordinal))) {
                    relationships.Add(new ExtendedRelationshipModel(
                        CreateRelationshipId(relationships),
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
                        "sharedStrings.xml",
                        isExternal: false));
                }

                foreach (var hyperlink in container.HyperlinkRelationships) {
                    relationships.Add(new ExtendedRelationshipModel(
                        hyperlink.Id,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                        hyperlink.Uri.ToString(),
                        hyperlink.IsExternal));
                }

                foreach (var external in container.ExternalRelationships) {
                    relationships.Add(new ExtendedRelationshipModel(
                        external.Id,
                        external.RelationshipType,
                        external.Uri.ToString(),
                        isExternal: true));
                }

                return relationships;
            }

            private static string CreateRelationshipId(IReadOnlyList<ExtendedRelationshipModel> relationships) {
                for (int i = 1; ; i++) {
                    string id = "rId" + InvariantNumberText.Get(i);
                    if (!relationships.Any(relationship => string.Equals(relationship.Id, id, StringComparison.Ordinal))) {
                        return id;
                    }
                }
            }
        }

        private static class FastWorkbookPackageWriter {
            internal static void Write(Stream destination, FastWorkbookPackageModel model, CancellationToken ct) {
                using (var archive = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true)) {
                    ct.ThrowIfCancellationRequested();
                    WriteContentTypesEntry(archive, model.HasStyles, model.HasSharedStrings, model.Worksheets.Count, model.Tables.Count);
                    ct.ThrowIfCancellationRequested();
                    WriteTextEntry(archive, "_rels/.rels",
                        "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                        "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>" +
                        "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>" +
                        "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>" +
                        "</Relationships>");
                    WriteCorePropertiesEntry(archive);
                    WriteAppPropertiesEntry(archive);
                    ct.ThrowIfCancellationRequested();
                    WriteWorkbookEntry(archive, model);
                    WriteWorkbookRelationshipsEntry(archive, model.Worksheets, model.HasStyles, model.HasSharedStrings);
                    if (model.Stylesheet != null) {
                        WriteTextEntry(archive, "xl/styles.xml", "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + model.Stylesheet.OuterXml);
                    }

                    if (model.HasSharedStrings && model.SharedStrings != null) {
                        WriteSharedStringsEntry(archive, model.SharedStrings);
                    }

                    foreach (var worksheet in model.Worksheets) {
                        ct.ThrowIfCancellationRequested();
                        WriteWorksheetEntry(archive, worksheet);
                        if (worksheet.HasRelationships) {
                            WriteWorksheetRelationshipsEntry(archive, worksheet);
                        }
                    }

                    for (int i = 0; i < model.Tables.Count; i++) {
                        ct.ThrowIfCancellationRequested();
                        WriteTextEntry(
                            archive,
                            "xl/tables/table" + InvariantNumberText.Get(i + 1) + ".xml",
                            "<?xml version=\"1.0\" encoding=\"utf-8\"?>" + model.Tables[i].OuterXml);
                    }
                }
            }
        }

        private sealed class FastWorkbookPackageModel {
            private FastWorkbookPackageModel(
                IReadOnlyList<FastWorksheetPackageModel> worksheets,
                Stylesheet? stylesheet,
                SharedStringTable? sharedStrings,
                IReadOnlyList<Table> tables,
                FileVersion? fileVersion,
                FileSharing? fileSharing,
                WorkbookProperties? workbookProperties,
                WorkbookProtection? workbookProtection,
                BookViews? bookViews,
                DefinedNames? definedNames,
                CalculationProperties? calculationProperties) {
                Worksheets = worksheets;
                Stylesheet = stylesheet;
                SharedStrings = sharedStrings;
                Tables = tables;
                FileVersion = fileVersion;
                FileSharing = fileSharing;
                WorkbookProperties = workbookProperties;
                WorkbookProtection = workbookProtection;
                BookViews = bookViews;
                DefinedNames = definedNames;
                CalculationProperties = calculationProperties;
            }

            internal IReadOnlyList<FastWorksheetPackageModel> Worksheets { get; }

            internal Stylesheet? Stylesheet { get; }

            internal bool HasStyles => Stylesheet != null;

            internal SharedStringTable? SharedStrings { get; }

            internal bool HasSharedStrings => SharedStrings != null && SharedStrings.Elements<SharedStringItem>().Any();

            internal IReadOnlyList<Table> Tables { get; }

            internal FileVersion? FileVersion { get; }

            internal FileSharing? FileSharing { get; }

            internal WorkbookProperties? WorkbookProperties { get; }

            internal WorkbookProtection? WorkbookProtection { get; }

            internal BookViews? BookViews { get; }

            internal DefinedNames? DefinedNames { get; }

            internal CalculationProperties? CalculationProperties { get; }

            internal static bool TryCreate(SpreadsheetDocument document, out FastWorkbookPackageModel model, out string? skipReason) {
                model = null!;
                skipReason = null;

                var workbookPart = document.WorkbookPart;
                if (workbookPart?.Workbook?.Sheets == null) {
                    skipReason = "Workbook is missing sheets.";
                    return false;
                }

                var sheets = workbookPart.Workbook.Sheets.OfType<Sheet>().ToList();
                if (sheets.Count == 0 || sheets.Any(sheet => sheet.Id == null)) {
                    skipReason = "Workbook has no sheets or has sheets without relationships.";
                    return false;
                }

                if (workbookPart.CalculationChainPart != null) {
                    skipReason = "Workbook contains a calculation chain part.";
                    return false;
                }

                var unsupportedWorkbookChild = workbookPart.Workbook.ChildElements
                    .FirstOrDefault(child => child is not DocumentFormat.OpenXml.Spreadsheet.FileVersion
                        && child is not DocumentFormat.OpenXml.Spreadsheet.FileSharing
                        && child is not DocumentFormat.OpenXml.Spreadsheet.WorkbookProperties
                        && child is not DocumentFormat.OpenXml.Spreadsheet.WorkbookProtection
                        && child is not DocumentFormat.OpenXml.Spreadsheet.BookViews
                        && child is not DocumentFormat.OpenXml.Spreadsheet.Sheets
                        && child is not DocumentFormat.OpenXml.Spreadsheet.DefinedNames
                        && child is not DocumentFormat.OpenXml.Spreadsheet.CalculationProperties);
                if (unsupportedWorkbookChild != null) {
                    skipReason = "Workbook contains unsupported workbook-level element '" + unsupportedWorkbookChild.LocalName + "'.";
                    return false;
                }

                foreach (var child in workbookPart.Workbook.ChildElements) {
                    if (child.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                        skipReason = "Workbook contains unknown Open XML elements.";
                        return false;
                    }
                }

                var definedNames = workbookPart.Workbook.GetFirstChild<DefinedNames>();
                if (workbookPart.GetPartsOfType<ThemePart>().Any()) {
                    skipReason = "Workbook contains a theme part.";
                    return false;
                }

                var worksheets = new List<FastWorksheetPackageModel>(sheets.Count);
                var tables = new List<Table>();
                int tableIndex = 1;
                for (int sheetIndex = 0; sheetIndex < sheets.Count; sheetIndex++) {
                    var sheet = sheets[sheetIndex];
                    if (workbookPart.GetPartById(sheet.Id!) is not WorksheetPart worksheetPart) {
                        skipReason = "Workbook sheet relationship does not target a worksheet part.";
                        return false;
                    }

                    var worksheet = worksheetPart.Worksheet;
                    if (worksheet == null) {
                        skipReason = "Worksheet part is missing worksheet XML.";
                        return false;
                    }

                    if (!CanWriteWorksheet(worksheetPart, worksheet, out skipReason)) {
                        return false;
                    }

                    var tablePartPaths = new Dictionary<string, string>(StringComparer.Ordinal);
                    foreach (var tableDefinition in worksheetPart.TableDefinitionParts) {
                        var table = tableDefinition.Table;
                        if (table == null) {
                            skipReason = "Worksheet table definition is missing table XML.";
                            return false;
                        }

                        tables.Add(table);
                        string relId = worksheetPart.GetIdOfPart(tableDefinition);
                        tablePartPaths[relId] = "../tables/table" + InvariantNumberText.Get(tableIndex) + ".xml";
                        tableIndex++;
                    }

                    var hyperlinkRelationships = worksheetPart.HyperlinkRelationships
                        .Select(relationship => new FastHyperlinkRelationshipModel(
                            relationship.Id,
                            relationship.Uri.ToString(),
                            relationship.IsExternal))
                        .ToList();

                    worksheets.Add(new FastWorksheetPackageModel(
                        sheet.Name?.Value ?? "Sheet" + InvariantNumberText.Get(sheetIndex + 1),
                        sheet.SheetId?.Value ?? (uint)(sheetIndex + 1),
                        GetSheetStateText(sheet),
                        "rId" + InvariantNumberText.Get(sheetIndex + 1),
                        "xl/worksheets/sheet" + InvariantNumberText.Get(sheetIndex + 1) + ".xml",
                        "xl/worksheets/_rels/sheet" + InvariantNumberText.Get(sheetIndex + 1) + ".xml.rels",
                        worksheet,
                        tablePartPaths,
                        hyperlinkRelationships));
                }

                var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable;
                if (sharedStrings != null
                    && sharedStrings.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                    skipReason = "Workbook shared strings contain unknown Open XML elements.";
                    return false;
                }

                model = new FastWorkbookPackageModel(
                    worksheets,
                    workbookPart.WorkbookStylesPart?.Stylesheet,
                    sharedStrings!,
                    tables,
                    workbookPart.Workbook.GetFirstChild<FileVersion>(),
                    workbookPart.Workbook.GetFirstChild<FileSharing>(),
                    workbookPart.Workbook.GetFirstChild<WorkbookProperties>(),
                    workbookPart.Workbook.GetFirstChild<WorkbookProtection>(),
                    workbookPart.Workbook.GetFirstChild<BookViews>(),
                    definedNames,
                    workbookPart.Workbook.GetFirstChild<CalculationProperties>());
                return true;
            }

            private static bool CanWriteWorksheet(WorksheetPart worksheetPart, Worksheet worksheet, out string? skipReason) {
                return CanWriteSimpleWorksheet(worksheetPart, worksheet, out skipReason);
            }

            private static string? GetSheetStateText(Sheet sheet) {
                if (sheet.State == null) {
                    return null;
                }

                if (sheet.State.Value == SheetStateValues.Hidden) {
                    return "hidden";
                }

                if (sheet.State.Value == SheetStateValues.VeryHidden) {
                    return "veryHidden";
                }

                if (sheet.State.Value == SheetStateValues.Visible) {
                    return "visible";
                }

                return sheet.State.InnerText;
            }
        }

        private sealed class FastWorksheetPackageModel {
            internal FastWorksheetPackageModel(
                string sheetName,
                uint sheetId,
                string? sheetState,
                string workbookRelationshipId,
                string worksheetPath,
                string relationshipsPath,
                Worksheet worksheet,
                IReadOnlyDictionary<string, string> tablePartPaths,
                IReadOnlyList<FastHyperlinkRelationshipModel> hyperlinkRelationships) {
                SheetName = sheetName;
                SheetId = sheetId;
                SheetState = sheetState;
                WorkbookRelationshipId = workbookRelationshipId;
                WorksheetPath = worksheetPath;
                RelationshipsPath = relationshipsPath;
                Worksheet = worksheet;
                TablePartPaths = tablePartPaths;
                HyperlinkRelationships = hyperlinkRelationships;
            }

            internal string SheetName { get; }

            internal uint SheetId { get; }

            internal string? SheetState { get; }

            internal string WorkbookRelationshipId { get; }

            internal string WorksheetPath { get; }

            internal string RelationshipsPath { get; }

            internal Worksheet Worksheet { get; }

            internal IReadOnlyDictionary<string, string> TablePartPaths { get; }

            internal IReadOnlyList<FastHyperlinkRelationshipModel> HyperlinkRelationships { get; }

            internal bool HasRelationships => TablePartPaths.Count > 0 || HyperlinkRelationships.Count > 0;
        }

        private sealed class FastHyperlinkRelationshipModel {
            internal FastHyperlinkRelationshipModel(string id, string target, bool isExternal) {
                Id = id;
                Target = target;
                IsExternal = isExternal;
            }

            internal string Id { get; }

            internal string Target { get; }

            internal bool IsExternal { get; }
        }

        private static bool CanWriteSimpleWorksheet(WorksheetPart worksheetPart, Worksheet worksheet, out string? skipReason, bool allowDrawings = false, bool allowPivotTables = false) {
            skipReason = null;

            if (worksheetPart.WorksheetCommentsPart != null) {
                skipReason = "Worksheet contains comments.";
                return false;
            }

            if (!allowDrawings && worksheetPart.DrawingsPart != null) {
                skipReason = "Worksheet contains drawings.";
                return false;
            }

            if (!allowPivotTables && worksheetPart.PivotTableParts.Any()) {
                skipReason = "Worksheet contains pivot tables.";
                return false;
            }

            if (worksheetPart.ExternalRelationships.Any()) {
                skipReason = "Worksheet contains external relationships.";
                return false;
            }

            foreach (var child in worksheet.ChildElements) {
                if (child is not SheetProperties
                    && child is not SheetDimension
                    && child is not SheetViews
                    && child is not SheetFormatProperties
                    && child is not Columns
                    && child is not SheetData
                    && child is not SheetCalculationProperties
                    && child is not SheetProtection
                    && child is not DocumentFormat.OpenXml.Spreadsheet.ProtectedRanges
                    && child is not Scenarios
                    && child is not AutoFilter
                    && child is not SortState
                    && child is not MergeCells
                    && child is not PhoneticProperties
                    && child is not Hyperlinks
                    && child is not DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting
                    && child is not DataValidations
                    && child is not PrintOptions
                    && child is not PageMargins
                    && child is not PageSetup
                    && child is not HeaderFooter
                    && child is not RowBreaks
                    && child is not ColumnBreaks
                    && child is not CellWatches
                    && child is not DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors
                    && (!allowDrawings || child is not DocumentFormat.OpenXml.Spreadsheet.Drawing)
                    && child is not TableParts) {
                    skipReason = "Worksheet contains unsupported element '" + child.LocalName + "'.";
                    return false;
                }

                if (child is not SheetData
                    && child.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                    skipReason = "Worksheet contains unknown Open XML elements.";
                    return false;
                }
            }

            var tableParts = worksheet.GetFirstChild<TableParts>();
            bool hasTableDefinitionParts = worksheetPart.TableDefinitionParts.Any();
            if (tableParts != null || hasTableDefinitionParts) {
                if (tableParts != null && worksheet.Elements<TableParts>().Skip(1).Any()) {
                    skipReason = "Worksheet contains multiple tableParts elements.";
                    return false;
                }

                var tableDefinitionParts = worksheetPart.TableDefinitionParts.ToList();
                var relationshipIds = new HashSet<string>(tableDefinitionParts.Select(worksheetPart.GetIdOfPart), StringComparer.Ordinal);
                var worksheetTablePartIds = tableParts == null
                    ? new List<string>()
                    : tableParts.Elements<TablePart>()
                        .Select(part => part.Id?.Value)
                        .Where(id => !string.IsNullOrEmpty(id))
                        .Select(id => id!)
                        .ToList();

                if (worksheetTablePartIds.Count != tableDefinitionParts.Count
                    || worksheetTablePartIds.Any(id => !relationshipIds.Contains(id))) {
                    skipReason = "Worksheet table relationships do not match tableParts entries.";
                    return false;
                }

                foreach (var tableDefinitionPart in tableDefinitionParts) {
                    var table = tableDefinitionPart.Table;
                    if (table == null
                        || table.Reference == null
                        || table.TableColumns == null
                        || table.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                        skipReason = "Worksheet contains unsupported table metadata.";
                        return false;
                    }
                }
            }

            var hyperlinks = worksheet.GetFirstChild<Hyperlinks>();
            bool hasHyperlinkRelationships = worksheetPart.HyperlinkRelationships.Any();
            if (hyperlinks != null || hasHyperlinkRelationships) {
                var hyperlinkRelationships = worksheetPart.HyperlinkRelationships.ToList();
                var hyperlinkIds = new HashSet<string>(hyperlinkRelationships.Select(relationship => relationship.Id), StringComparer.Ordinal);
                if (hyperlinks != null) {
                    foreach (var hyperlink in hyperlinks.Elements<Hyperlink>()) {
                        string? relationshipId = hyperlink.Id?.Value;
                        if (!string.IsNullOrEmpty(relationshipId) && !hyperlinkIds.Contains(relationshipId!)) {
                            skipReason = "Worksheet hyperlink relationships do not match hyperlink entries.";
                            return false;
                        }
                    }
                }
            }

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return true;
            }

            foreach (var sheetDataChild in sheetData.ChildElements) {
                if (sheetDataChild is not Row) {
                    skipReason = sheetDataChild is DocumentFormat.OpenXml.OpenXmlUnknownElement
                        ? "Worksheet contains unknown Open XML elements."
                        : "Worksheet contains sheetData children outside the simple writer surface.";
                    return false;
                }
            }

            foreach (var row in sheetData.Elements<Row>()) {
                if (!IsSimpleRow(row)) {
                    skipReason = "Worksheet contains row formatting outside the simple writer surface.";
                    return false;
                }

                foreach (var rowChild in row.ChildElements) {
                    if (rowChild is DocumentFormat.OpenXml.OpenXmlUnknownElement) {
                        skipReason = "Worksheet contains unknown Open XML elements.";
                        return false;
                    }

                    if (rowChild is not Cell cell) {
                        skipReason = "Worksheet contains row children outside the simple writer surface.";
                        return false;
                    }

                    foreach (var cellChild in cell.ChildElements) {
                        if (cellChild is DocumentFormat.OpenXml.OpenXmlUnknownElement) {
                            skipReason = "Worksheet contains unknown Open XML elements.";
                            return false;
                        }
                    }

                    if (cell.InlineString != null) {
                        if (cell.InlineString.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                            skipReason = "Worksheet inline strings contain unknown Open XML elements.";
                            return false;
                        }
                    }

                    if (cell.CellFormula != null
                        && cell.CellFormula.Descendants<DocumentFormat.OpenXml.OpenXmlUnknownElement>().Any()) {
                        skipReason = "Worksheet contains formula metadata outside the simple writer surface.";
                        return false;
                    }

                    var dataType = cell.DataType?.Value;
                    if (dataType != null
                        && dataType != CellValues.Number
                        && dataType != CellValues.SharedString
                        && dataType != CellValues.InlineString
                        && dataType != CellValues.String
                        && dataType != CellValues.Boolean) {
                        skipReason = "Worksheet contains unsupported cell data type '" + dataType.Value.ToString() + "'.";
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool IsSimpleRow(Row row) {
            foreach (var attribute in row.GetAttributes()) {
                if (!string.Equals(attribute.LocalName, "r", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "hidden", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "ht", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "customHeight", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "outlineLevel", StringComparison.Ordinal)
                    && !string.Equals(attribute.LocalName, "collapsed", StringComparison.Ordinal)) {
                    return false;
                }
            }

            return row.CustomFormat?.Value != true && row.StyleIndex == null;
        }

        private static void WriteContentTypesEntry(ZipArchive archive, bool hasStyles, bool hasSharedStrings, int worksheetCount, int tableCount) {
            var builder = new System.Text.StringBuilder(512 + worksheetCount * 160 + tableCount * 160);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            builder.Append("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>");
            builder.Append("<Default Extension=\"xml\" ContentType=\"application/xml\"/>");
            builder.Append("<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>");
            builder.Append("<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>");
            builder.Append("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>");
            for (int i = 1; i <= worksheetCount; i++) {
                builder.Append("<Override PartName=\"/xl/worksheets/sheet");
                AppendInvariant(builder, i);
                builder.Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
            }

            if (hasStyles) {
                builder.Append("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>");
            }

            if (hasSharedStrings) {
                builder.Append("<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>");
            }

            for (int i = 1; i <= tableCount; i++) {
                builder.Append("<Override PartName=\"/xl/tables/table");
                AppendInvariant(builder, i);
                builder.Append(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>");
            }

            builder.Append("</Types>");
            WriteTextEntry(archive, "[Content_Types].xml", builder.ToString());
        }

        private static void WriteWorkbookEntry(ZipArchive archive, FastWorkbookPackageModel model) {
            var entry = archive.CreateEntry("xl/workbook.xml", CompressionLevel.Fastest);
            var worksheets = model.Worksheets;
            using var stream = entry.Open();
            using var writer = CreateFastXmlWriter(stream);
            writer.WriteStartDocument();
            writer.WriteStartElement("workbook", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            writer.WriteAttributeString("xmlns", "r", null, "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            if (model.FileVersion != null) {
                model.FileVersion.WriteTo(writer);
            }

            if (model.FileSharing != null) {
                model.FileSharing.WriteTo(writer);
            }

            if (model.WorkbookProperties != null) {
                model.WorkbookProperties.WriteTo(writer);
            }

            if (model.WorkbookProtection != null) {
                model.WorkbookProtection.WriteTo(writer);
            }

            if (model.BookViews != null) {
                model.BookViews.WriteTo(writer);
            }

            writer.WriteStartElement("sheets");
            foreach (var worksheet in worksheets) {
                writer.WriteStartElement("sheet");
                writer.WriteAttributeString("name", worksheet.SheetName);
                writer.WriteAttributeString("sheetId", InvariantNumberText.Get(worksheet.SheetId));
                writer.WriteAttributeString("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", worksheet.WorkbookRelationshipId);
                if (!string.IsNullOrEmpty(worksheet.SheetState)) {
                    writer.WriteAttributeString("state", worksheet.SheetState);
                }

                writer.WriteEndElement();
            }

            writer.WriteEndElement();
            if (model.DefinedNames != null) {
                model.DefinedNames.WriteTo(writer);
            }

            if (model.CalculationProperties != null) {
                model.CalculationProperties.WriteTo(writer);
            }

            writer.WriteEndElement();
        }

        private static void WriteWorkbookRelationshipsEntry(ZipArchive archive, IReadOnlyList<FastWorksheetPackageModel> worksheets, bool hasStyles, bool hasSharedStrings) {
            var builder = new System.Text.StringBuilder(384 + worksheets.Count * 180);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var worksheet in worksheets) {
                builder.Append("<Relationship Id=\"");
                AppendXmlEscaped(builder, worksheet.WorkbookRelationshipId);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"");
                AppendXmlEscaped(builder, worksheet.WorksheetPath.Substring("xl/".Length));
                builder.Append("\"/>");
            }

            if (hasStyles) {
                builder.Append("<Relationship Id=\"rId");
                AppendInvariant(builder, worksheets.Count + 1);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>");
            }

            if (hasSharedStrings) {
                builder.Append("<Relationship Id=\"rId");
                AppendInvariant(builder, worksheets.Count + (hasStyles ? 2 : 1));
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>");
            }

            builder.Append("</Relationships>");
            WriteTextEntry(archive, "xl/_rels/workbook.xml.rels", builder.ToString());
        }

        private static void WriteCorePropertiesEntry(ZipArchive archive) {
            WriteTextEntry(archive, "docProps/core.xml",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" " +
                "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" " +
                "xmlns:dcterms=\"http://purl.org/dc/terms/\" " +
                "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" " +
                "xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"/>");
        }

        private static void WriteAppPropertiesEntry(ZipArchive archive) {
            WriteTextEntry(archive, "docProps/app.xml",
                "<?xml version=\"1.0\" encoding=\"utf-8\"?>" +
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\" " +
                "xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">" +
                "<Application>OfficeIMO.Excel</Application>" +
                "</Properties>");
        }

        private static void WriteSharedStringsEntry(ZipArchive archive, SharedStringTable sharedStrings) {
            WriteOpenXmlElementEntry(archive, "xl/sharedStrings.xml", sharedStrings);
        }

        private static void WriteOpenXmlElementEntry(ZipArchive archive, string path, OpenXmlElement element) {
            var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = CreateFastXmlWriter(stream);
            writer.WriteStartDocument();
            element.WriteTo(writer);
        }

        private static void WriteWorksheetRelationshipsEntry(ZipArchive archive, FastWorksheetPackageModel worksheet) {
            var builder = new System.Text.StringBuilder(160 + worksheet.TablePartPaths.Count * 180 + worksheet.HyperlinkRelationships.Count * 220);
            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var item in worksheet.TablePartPaths.OrderBy(static item => item.Key, StringComparer.Ordinal)) {
                builder.Append("<Relationship Id=\"");
                AppendXmlEscaped(builder, item.Key);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"");
                AppendXmlEscaped(builder, item.Value);
                builder.Append("\"/>");
            }

            foreach (var relationship in worksheet.HyperlinkRelationships.OrderBy(static item => item.Id, StringComparer.Ordinal)) {
                builder.Append("<Relationship Id=\"");
                AppendXmlEscaped(builder, relationship.Id);
                builder.Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"");
                AppendXmlEscaped(builder, relationship.Target);
                builder.Append('"');
                if (relationship.IsExternal) {
                    builder.Append(" TargetMode=\"External\"");
                }

                builder.Append("/>");
            }

            builder.Append("</Relationships>");
            WriteTextEntry(archive, worksheet.RelationshipsPath, builder.ToString());
        }

        private static void WriteWorksheetEntry(ZipArchive archive, FastWorksheetPackageModel model) {
            var entry = archive.CreateEntry(model.WorksheetPath, CompressionLevel.Fastest);
            var worksheet = model.Worksheet;
            string dimension = worksheet.SheetDimension?.Reference?.Value ?? ExcelSheet.ComputeSheetDimensionReference(worksheet);
            var builder = new System.Text.StringBuilder(4096);
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream, Utf8NoBom);

            builder.Append("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            builder.Append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"");
            if (model.HasRelationships) {
                builder.Append(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
            }

            builder.Append(">");
            WriteBuilderAndClear(writer, builder);
            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetProperties>());

            builder.Append("<dimension ref=\"");
            AppendXmlEscaped(builder, dimension);
            builder.Append("\"/>");
            WriteBuilderAndClear(writer, builder);

            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetViews>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetFormatProperties>());

            var columns = worksheet.GetFirstChild<Columns>();
            if (columns != null) {
                AppendColumns(builder, columns);
                WriteBuilderAndClear(writer, builder);
            }

            writer.Write("<sheetData>");

            var sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData != null) {
                foreach (var row in sheetData.Elements<Row>()) {
                    AppendSimpleRowStart(builder, row);

                    foreach (var cell in row.Elements<Cell>()) {
                        AppendSimpleCell(builder, cell);
                    }

                    builder.Append("</row>");
                    WriteBuilderAndClear(writer, builder);
                }
            }

            writer.Write("</sheetData>");
            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetCalculationProperties>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<SheetProtection>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.ProtectedRanges>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<Scenarios>());

            var autoFilter = worksheet.GetFirstChild<AutoFilter>();
            if (autoFilter != null) {
                writer.Write(autoFilter.OuterXml);
            }

            WriteOptionalElement(writer, worksheet.GetFirstChild<SortState>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<MergeCells>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<PhoneticProperties>());
            WriteOptionalElements<DocumentFormat.OpenXml.Spreadsheet.ConditionalFormatting>(writer, worksheet);
            WriteOptionalElement(writer, worksheet.GetFirstChild<DataValidations>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<Hyperlinks>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<PrintOptions>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<PageMargins>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<PageSetup>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<HeaderFooter>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<RowBreaks>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<ColumnBreaks>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<CellWatches>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.IgnoredErrors>());
            WriteOptionalElement(writer, worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Drawing>());

            var tableParts = worksheet.GetFirstChild<TableParts>();
            if (tableParts != null && model.TablePartPaths.Count > 0) {
                builder.Append("<tableParts count=\"");
                AppendInvariant(builder, model.TablePartPaths.Count);
                builder.Append("\">");
                foreach (var tablePart in tableParts.Elements<TablePart>()) {
                    string? id = tablePart.Id?.Value;
                    if (id == null || !model.TablePartPaths.ContainsKey(id)) {
                        continue;
                    }

                    builder.Append("<tablePart r:id=\"");
                    AppendXmlEscaped(builder, id);
                    builder.Append("\"/>");
                }

                builder.Append("</tableParts>");
                WriteBuilderAndClear(writer, builder);
            }

            writer.Write("</worksheet>");
        }

        private static void WriteBuilderAndClear(StreamWriter writer, System.Text.StringBuilder builder) {
            if (builder.Length == 0) {
                return;
            }

#if NET6_0_OR_GREATER
            writer.Write(builder);
#else
            writer.Write(builder.ToString());
#endif
            builder.Clear();
            if (builder.Capacity > 65536) {
                builder.Capacity = 4096;
            }
        }

        private static void WriteOptionalElement(StreamWriter writer, OpenXmlElement? element) {
            if (element != null) {
                writer.Write(element.OuterXml);
            }
        }

        private static void WriteOptionalElements<TElement>(StreamWriter writer, OpenXmlElement parent)
            where TElement : OpenXmlElement {
            foreach (var element in parent.Elements<TElement>()) {
                writer.Write(element.OuterXml);
            }
        }

        private static void AppendColumns(System.Text.StringBuilder builder, Columns columns) {
            builder.Append("<cols>");
            foreach (var column in columns.Elements<Column>()) {
                builder.Append("<col");
                AppendUIntAttribute(builder, "min", column.Min);
                AppendUIntAttribute(builder, "max", column.Max);
                if (column.Width != null) {
                    builder.Append(" width=\"");
                    builder.Append(column.Width.Value.ToString(CultureInfo.InvariantCulture));
                    builder.Append('"');
                }

                AppendBooleanAttribute(builder, "bestFit", column.BestFit);
                AppendBooleanAttribute(builder, "customWidth", column.CustomWidth);
                AppendBooleanAttribute(builder, "hidden", column.Hidden);
                AppendUIntAttribute(builder, "style", column.Style);
                AppendByteAttribute(builder, "outlineLevel", column.OutlineLevel);
                AppendBooleanAttribute(builder, "collapsed", column.Collapsed);
                AppendBooleanAttribute(builder, "phonetic", column.Phonetic);
                builder.Append("/>");
            }

            builder.Append("</cols>");
        }

        private static void AppendSimpleRowStart(System.Text.StringBuilder builder, Row row) {
            builder.Append("<row");
            AppendUIntAttribute(builder, "r", row.RowIndex);
            AppendBooleanAttribute(builder, "hidden", row.Hidden);
            if (row.Height != null) {
                builder.Append(" ht=\"");
                builder.Append(row.Height.Value.ToString(CultureInfo.InvariantCulture));
                builder.Append('"');
            }

            AppendBooleanAttribute(builder, "customHeight", row.CustomHeight);
            AppendByteAttribute(builder, "outlineLevel", row.OutlineLevel);
            AppendBooleanAttribute(builder, "collapsed", row.Collapsed);
            builder.Append('>');
        }

        private static void AppendSimpleCell(System.Text.StringBuilder builder, Cell cell) {
            string? text = cell.CellValue?.Text;
            var dataType = cell.DataType?.Value;

            builder.Append("<c");
            if (cell.CellReference != null) {
                builder.Append(" r=\"");
                AppendXmlEscaped(builder, cell.CellReference.Value ?? string.Empty);
                builder.Append('"');
            }

            if (cell.StyleIndex != null) {
                builder.Append(" s=\"");
                AppendInvariant(builder, cell.StyleIndex.Value);
                builder.Append('"');
            }

            if (dataType == CellValues.Number) {
                builder.Append(" t=\"n\"");
            } else if (dataType == CellValues.SharedString) {
                builder.Append(" t=\"s\"");
            } else if (dataType == CellValues.InlineString || cell.InlineString != null) {
                builder.Append(" t=\"inlineStr\"");
            } else if (dataType == CellValues.String) {
                builder.Append(" t=\"str\"");
            } else if (dataType == CellValues.Boolean) {
                builder.Append(" t=\"b\"");
            }

            builder.Append('>');
            if (cell.CellFormula != null) {
                AppendCellFormula(builder, cell.CellFormula);
            }

            if (cell.InlineString != null) {
                builder.Append(cell.InlineString.OuterXml);
                builder.Append("</c>");
                return;
            }

            if (cell.CellValue != null) {
                string valueText = text ?? string.Empty;
                if (valueText.Length == 0) {
                    builder.Append("<v/>");
                } else {
                    builder.Append("<v>");
                    AppendXmlEscaped(builder, valueText);
                    builder.Append("</v>");
                }
            }

            builder.Append("</c>");
        }

        private static void AppendCellFormula(System.Text.StringBuilder builder, CellFormula formula) {
            if (formula.HasAttributes) {
                builder.Append(formula.OuterXml);
                return;
            }

            builder.Append("<f>");
            AppendXmlEscaped(builder, formula.Text ?? string.Empty);
            builder.Append("</f>");
        }

        private static void WriteTextEntry(ZipArchive archive, string path, string text) {
            var entry = archive.CreateEntry(path, CompressionLevel.Fastest);
            using var stream = entry.Open();
            using var writer = new StreamWriter(stream, Utf8NoBom);
            writer.Write(text);
        }

        private static void AppendUIntAttribute(System.Text.StringBuilder builder, string name, UInt32Value? value) {
            if (value == null) {
                return;
            }

            builder.Append(' ');
            builder.Append(name);
            builder.Append("=\"");
            AppendInvariant(builder, value.Value);
            builder.Append('"');
        }

        private static void AppendByteAttribute(System.Text.StringBuilder builder, string name, ByteValue? value) {
            if (value == null) {
                return;
            }

            builder.Append(' ');
            builder.Append(name);
            builder.Append("=\"");
            AppendInvariant(builder, value.Value);
            builder.Append('"');
        }

        private static void AppendBooleanAttribute(System.Text.StringBuilder builder, string name, BooleanValue? value) {
            if (value == null) {
                return;
            }

            builder.Append(' ');
            builder.Append(name);
            builder.Append("=\"");
            builder.Append(value.Value ? '1' : '0');
            builder.Append('"');
        }

        private static void AppendInvariant(System.Text.StringBuilder builder, int value)
            => builder.Append(InvariantNumberText.Get(value));

        private static void AppendInvariant(System.Text.StringBuilder builder, uint value)
            => builder.Append(InvariantNumberText.Get(value));

        private static XmlWriter CreateFastXmlWriter(Stream stream) =>
            XmlWriter.Create(stream, new XmlWriterSettings {
                Encoding = Utf8NoBom,
                CloseOutput = false,
                Indent = false,
                OmitXmlDeclaration = false
            });

        private static void AppendXmlEscaped(System.Text.StringBuilder builder, string text) {
            for (int i = 0; i < text.Length; i++) {
                char ch = text[i];
                switch (ch) {
                    case '&':
                        builder.Append("&amp;");
                        break;
                    case '<':
                        builder.Append("&lt;");
                        break;
                    case '>':
                        builder.Append("&gt;");
                        break;
                    case '"':
                        builder.Append("&quot;");
                        break;
                    case '\'':
                        builder.Append("&apos;");
                        break;
                    default:
                        builder.Append(ch);
                        break;
                }
            }
        }
    }
}
