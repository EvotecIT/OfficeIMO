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
                bool useMixedDirectWorksheetEntries = CanUseMixedDirectWorksheetEntries(model);
                if ((useMixedDirectWorksheetEntries || CanUseDirectWorksheetEntries(model))
                    && DirectDataSetWorkbookWriter.TryCreateExtendedWritePlan(
                        model.DirectDataSetModel!,
                        ct,
                        out var candidateDirectWritePlan,
                        disableSharedStrings: useMixedDirectWorksheetEntries)) {
                    directWritePlan = candidateDirectWritePlan;
                }
                ReportTiming("Save.ExtendedPackage.CreateDirectWritePlan");

                WriteExtendedContentTypesEntry(archive, model.Parts, directWritePlan != null, directWritePlan?.HasSharedStrings == true);
                WriteExtendedRelationshipsEntry(archive, "_rels/.rels", model.PackageRelationships);
                WriteTextEntry(archive, "docProps/core.xml", model.CorePropertiesXml);
                WriteTextEntry(archive, "docProps/app.xml", model.AppPropertiesXml);
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
                    } else if (part.CopyRawPart) {
                        WriteRawPartEntry(archive, part.Path, part.Part);
                        ReportTiming("Save.ExtendedPackage.WriteRawPart");
                    } else if (part.Part is WorksheetPart worksheetPart
                        && CanWriteSimpleWorksheet(worksheetPart, worksheetPart.Worksheet!, out _, allowDrawings: true, allowPivotTables: true)) {
                        var tablePartIds = worksheetPart.TableDefinitionParts
                            .Select(worksheetPart.GetIdOfPart)
                            .ToDictionary(static id => id, static id => string.Empty, StringComparer.Ordinal);
                        var hyperlinkRelationships = worksheetPart.HyperlinkRelationships
                            .Select(static relationship => new FastHyperlinkRelationshipModel(
                                relationship.Id,
                                relationship.Uri.ToString(),
                                relationship.IsExternal))
                            .ToList();
                        var worksheetModel = new FastWorksheetPackageModel(
                            string.Empty,
                            0U,
                            null,
                            string.Empty,
                            part.Path,
                            GetRelationshipsPath(part.Path),
                            worksheetPart.Worksheet!,
                            tablePartIds,
                            hyperlinkRelationships);
                        WriteWorksheetEntry(archive, worksheetModel);
                        ReportTiming("Save.ExtendedPackage.WriteSimpleWorksheet");
                    } else {
                        WriteOpenXmlElementEntry(archive, part.Path, part.RootElement!);
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

                if (model.DirectWorksheetModels.Count != worksheetPartCount
                    && !CanUseMixedDirectWorksheetEntries(model)) {
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

    }
}
