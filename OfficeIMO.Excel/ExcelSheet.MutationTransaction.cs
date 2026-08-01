using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing.Internal;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal ExcelMutationResult ApplyTransactionalMutation(
            Action<CancellationToken> operation,
            int affectedCells,
            ExcelMutationPlanOptions options,
            CancellationToken cancellationToken) =>
            ApplyTransactionalMutation(token => {
                operation(token);
                return affectedCells;
            }, options, cancellationToken);

        internal ExcelMutationResult ApplyTransactionalMutation(
            Func<CancellationToken, int> operation,
            ExcelMutationPlanOptions options,
            CancellationToken cancellationToken) {
            ExcelMutationResult? result = null;
            Batch(_ => {
                cancellationToken.ThrowIfCancellationRequested();
                EnsureWorksheetCapturedByMutationPlanIsActive();
                var snapshot = PackageMutationSnapshot.Capture(_excelDocument.WorkbookPartRoot, options.MaximumSnapshotCharacters);
                try {
                    int affectedCells = operation(cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();
                    WorksheetRoot.Save();
                    MarkRequiresSavePreparation();
                    IReadOnlyList<ExcelMutationDiagnostic> diagnostics =
                        _excelDocument.GetMutationDiagnostics(options.MaximumDiagnostics, cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();
                    result = new ExcelMutationResult(
                        affectedCells,
                        diagnostics);
                } catch {
                    snapshot.Restore();
                    ResetMutationCaches();
                    throw;
                }
            });
            return result!;
        }

        private void ResetMutationCaches() {
            _sheetDataCache = null;
            _lastAccessedRow = null;
            _lastAccessedCell = null;
            _lastAccessedRowIndex = 0;
            _lastAccessedCellRowIndex = 0;
            _lastAccessedCellColumnIndex = 0;
            ClearHeaderCache();
        }

        /// <summary>Reads a payload-only rollback part without exceeding the unconsumed snapshot budget.</summary>
        internal static byte[] ReadMutationSnapshotPayload(
            Stream source,
            long remainingCharacters,
            long maximumCharacters) {
            if (remainingCharacters < 1) {
                throw new InvalidOperationException(
                    $"Transactional snapshot exceeds MaximumSnapshotCharacters ({maximumCharacters}).");
            }
            try {
                return OfficeStreamReader.ReadAllBytes(source, remainingCharacters);
            } catch (InvalidDataException exception) {
                throw new InvalidOperationException(
                    $"Transactional snapshot exceeds MaximumSnapshotCharacters ({maximumCharacters}).",
                    exception);
            }
        }

        private sealed class PackageMutationSnapshot {
            private readonly List<Action> _restore = new List<Action>();

            internal static PackageMutationSnapshot Capture(WorkbookPart workbookPart, long maximumCharacters) {
                var snapshot = new PackageMutationSnapshot();
                long characters = 0;
                void AddRoot<T>(T? root, Action<T> restore) where T : OpenXmlPartRootElement {
                    if (root == null) return;
                    string xml = root.OuterXml;
                    characters = checked(characters + xml.Length);
                    if (characters > maximumCharacters) {
                        throw new InvalidOperationException($"Transactional snapshot exceeds MaximumSnapshotCharacters ({maximumCharacters}).");
                    }
                    T clone = (T)root.CloneNode(true);
                    snapshot._restore.Add(() => restore(clone));
                }

                TPart RestorePartRelationship<TPart>(
                    OpenXmlPartContainer parent,
                    string relationshipId) where TPart : OpenXmlPart, IFixedContentTypePart {
                    foreach (IdPartPair pair in parent.Parts) {
                        if (!string.Equals(pair.RelationshipId, relationshipId, StringComparison.Ordinal)) continue;
                        return pair.OpenXmlPart as TPart
                            ?? throw new InvalidOperationException(
                                $"Relationship '{relationshipId}' was restored with an incompatible package-part type.");
                    }
                    return parent.AddNewPart<TPart>(relationshipId);
                }

                void AddPartRoot<TPart, TRoot>(
                    OpenXmlPartContainer parent,
                    TPart part,
                    TRoot? root,
                    Action<TPart, TRoot> restoreRoot)
                    where TPart : OpenXmlPart, IFixedContentTypePart
                    where TRoot : OpenXmlPartRootElement {
                    if (root == null) return;
                    string relationshipId = parent.GetIdOfPart(part);
                    string xml = root.OuterXml;
                    characters = checked(characters + xml.Length);
                    if (characters > maximumCharacters) {
                        throw new InvalidOperationException($"Transactional snapshot exceeds MaximumSnapshotCharacters ({maximumCharacters}).");
                    }
                    TRoot clone = (TRoot)root.CloneNode(true);
                    snapshot._restore.Add(() => restoreRoot(
                        RestorePartRelationship<TPart>(parent, relationshipId),
                        (TRoot)clone.CloneNode(true)));
                }

                void AddPartPayload<TPart>(
                    OpenXmlPartContainer parent,
                    TPart part) where TPart : OpenXmlPart, IFixedContentTypePart {
                    string relationshipId = parent.GetIdOfPart(part);
                    using Stream source = part.GetStream(FileMode.Open, FileAccess.Read);
                    byte[] bytes = ReadMutationSnapshotPayload(
                        source,
                        maximumCharacters - characters,
                        maximumCharacters);
                    characters = checked(characters + bytes.Length);
                    if (characters > maximumCharacters) {
                        throw new InvalidOperationException($"Transactional snapshot exceeds MaximumSnapshotCharacters ({maximumCharacters}).");
                    }
                    snapshot._restore.Add(() => {
                        TPart restoredPart = RestorePartRelationship<TPart>(parent, relationshipId);
                        using var restore = new MemoryStream(bytes, writable: false);
                        restoredPart.FeedData(restore);
                    });
                }

                AddRoot(workbookPart.Workbook, value => workbookPart.Workbook = value);
                AddRoot(workbookPart.ConnectionsPart?.Connections, value => workbookPart.ConnectionsPart!.Connections = value);
                AddRoot(workbookPart.WorkbookStylesPart?.Stylesheet, value => workbookPart.WorkbookStylesPart!.Stylesheet = value);
                AddRoot(workbookPart.SharedStringTablePart?.SharedStringTable, value => workbookPart.SharedStringTablePart!.SharedStringTable = value);
                foreach (SlicerCachePart part in workbookPart.SlicerCacheParts) {
                    AddPartRoot(
                        workbookPart,
                        part,
                        part.SlicerCacheDefinition,
                        (restoredPart, value) => restoredPart.SlicerCacheDefinition = value);
                }
                foreach (TimeLineCachePart part in workbookPart.TimeLineCacheParts) {
                    AddPartRoot(
                        workbookPart,
                        part,
                        part.TimelineCacheDefinition,
                        (restoredPart, value) => restoredPart.TimelineCacheDefinition = value);
                }
                CalculationChainPart? calculationChainPart = workbookPart.CalculationChainPart;
                if (calculationChainPart?.CalculationChain != null) {
                    string xml = calculationChainPart.CalculationChain.OuterXml;
                    characters = checked(characters + xml.Length);
                    if (characters > maximumCharacters) {
                        throw new InvalidOperationException($"Transactional snapshot exceeds MaximumSnapshotCharacters ({maximumCharacters}).");
                    }
                    CalculationChain clone = (CalculationChain)calculationChainPart.CalculationChain.CloneNode(true);
                    snapshot._restore.Add(() => {
                        CalculationChainPart restoredPart = workbookPart.CalculationChainPart
                            ?? workbookPart.AddNewPart<CalculationChainPart>();
                        restoredPart.CalculationChain = (CalculationChain)clone.CloneNode(true);
                    });
                } else {
                    snapshot._restore.Add(() => {
                        CalculationChainPart? createdPart = workbookPart.CalculationChainPart;
                        if (createdPart != null) workbookPart.DeletePart(createdPart);
                    });
                }
                void AddDrawingRoots(DrawingsPart? drawingsPart) {
                    if (drawingsPart == null) return;
                    AddRoot(drawingsPart.WorksheetDrawing, value => drawingsPart.WorksheetDrawing = value);
                    foreach (ChartPart part in drawingsPart.ChartParts) {
                        AddRoot(part.ChartSpace, value => part.ChartSpace = value);
                    }
                    foreach (ExtendedChartPart part in drawingsPart.ExtendedChartParts) {
                        AddRoot(part.ChartSpace, value => part.ChartSpace = value);
                    }
                }

                foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts) {
                    AddRoot(worksheetPart.Worksheet, value => worksheetPart.Worksheet = value);
                    var hyperlinkRelationships = worksheetPart.HyperlinkRelationships
                        .Select(relationship => (
                            relationship.Id,
                            relationship.Uri,
                            relationship.IsExternal))
                        .ToArray();
                    characters = checked(characters + hyperlinkRelationships.Sum(relationship =>
                        relationship.Id.Length + relationship.Uri.OriginalString.Length));
                    if (characters > maximumCharacters) {
                        throw new InvalidOperationException($"Transactional snapshot exceeds MaximumSnapshotCharacters ({maximumCharacters}).");
                    }
                    snapshot._restore.Add(() => {
                        var baselineIds = new HashSet<string>(
                            hyperlinkRelationships.Select(relationship => relationship.Id),
                            StringComparer.Ordinal);
                        foreach (HyperlinkRelationship relationship in worksheetPart.HyperlinkRelationships.ToList()) {
                            if (!baselineIds.Contains(relationship.Id)) {
                                worksheetPart.DeleteReferenceRelationship(relationship);
                            }
                        }
                        foreach (var baseline in hyperlinkRelationships) {
                            HyperlinkRelationship? current = worksheetPart.HyperlinkRelationships
                                .FirstOrDefault(relationship => string.Equals(relationship.Id, baseline.Id, StringComparison.Ordinal));
                            if (current != null
                                && current.Uri == baseline.Uri
                                && current.IsExternal == baseline.IsExternal) continue;
                            if (current != null) worksheetPart.DeleteReferenceRelationship(current);
                            worksheetPart.AddHyperlinkRelationship(
                                baseline.Uri,
                                baseline.IsExternal,
                                baseline.Id);
                        }
                    });
                    WorksheetCommentsPart? commentsPart = worksheetPart.WorksheetCommentsPart;
                    if (commentsPart != null) {
                        if (commentsPart.Comments == null) AddPartPayload(worksheetPart, commentsPart);
                        else {
                            AddPartRoot(
                                worksheetPart,
                                commentsPart,
                                commentsPart.Comments,
                                (part, value) => part.Comments = value);
                        }
                    }
                    foreach (WorksheetThreadedCommentsPart part in worksheetPart.WorksheetThreadedCommentsParts) {
                        if (part.ThreadedComments == null) AddPartPayload(worksheetPart, part);
                        else {
                            AddPartRoot(
                                worksheetPart,
                                part,
                                part.ThreadedComments,
                                (restoredPart, value) => restoredPart.ThreadedComments = value);
                        }
                    }
                    foreach (NamedSheetViewsPart part in worksheetPart.NamedSheetViewsParts) {
                        if (part.NamedSheetViews == null) AddPartPayload(worksheetPart, part);
                        else {
                            AddPartRoot(
                                worksheetPart,
                                part,
                                part.NamedSheetViews,
                                (restoredPart, value) => restoredPart.NamedSheetViews = value);
                        }
                    }
                    foreach (SlicersPart part in worksheetPart.SlicersParts) {
                        AddPartRoot(
                            worksheetPart,
                            part,
                            part.Slicers,
                            (restoredPart, value) => restoredPart.Slicers = value);
                    }
                    foreach (TimeLinePart part in worksheetPart.TimeLineParts) {
                        AddPartRoot(
                            worksheetPart,
                            part,
                            part.Timelines,
                            (restoredPart, value) => restoredPart.Timelines = value);
                    }
                    AddDrawingRoots(worksheetPart.DrawingsPart);
                    foreach (TableDefinitionPart part in worksheetPart.TableDefinitionParts) {
                        AddPartRoot(
                            worksheetPart,
                            part,
                            part.Table,
                            (restoredPart, value) => restoredPart.Table = value);
                        foreach (QueryTablePart queryPart in part.QueryTableParts) {
                            AddPartRoot(
                                part,
                                queryPart,
                                queryPart.QueryTable,
                                (restoredPart, value) => restoredPart.QueryTable = value);
                        }
                    }
                    foreach (QueryTablePart part in worksheetPart.QueryTableParts) {
                        AddRoot(part.QueryTable, value => part.QueryTable = value);
                    }
                    foreach (PivotTablePart part in worksheetPart.PivotTableParts) {
                        AddRoot(part.PivotTableDefinition, value => part.PivotTableDefinition = value);
                    }
                }
                foreach (ChartsheetPart chartsheetPart in workbookPart.ChartsheetParts) {
                    AddDrawingRoots(chartsheetPart.DrawingsPart);
                }
                foreach (PivotTableCacheDefinitionPart part in workbookPart.PivotTableCacheDefinitionParts) {
                    AddRoot(part.PivotCacheDefinition, value => part.PivotCacheDefinition = value);
                    PivotTableCacheRecordsPart? records = part.PivotTableCacheRecordsPart;
                    AddRoot(records?.PivotCacheRecords, value => records!.PivotCacheRecords = value);
                }
                foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts) {
                    foreach (VmlDrawingPart part in worksheetPart.VmlDrawingParts) {
                        AddPartPayload(worksheetPart, part);
                    }
                }
                var relationshipBaselines = new Dictionary<OpenXmlPartContainer, HashSet<string>>();
                var pending = new Stack<OpenXmlPartContainer>();
                pending.Push(workbookPart);
                while (pending.Count > 0) {
                    OpenXmlPartContainer parent = pending.Pop();
                    if (relationshipBaselines.ContainsKey(parent)) continue;
                    IdPartPair[] parts = parent.Parts.ToArray();
                    relationshipBaselines[parent] = new HashSet<string>(parts.Select(pair => pair.RelationshipId), StringComparer.Ordinal);
                    foreach (IdPartPair pair in parts) pending.Push(pair.OpenXmlPart);
                }
                snapshot._restore.Add(() => {
                    foreach (KeyValuePair<OpenXmlPartContainer, HashSet<string>> baseline in relationshipBaselines) {
                        List<IdPartPair> addedParts;
                        try {
                            addedParts = baseline.Key.Parts
                                .Where(pair => !baseline.Value.Contains(pair.RelationshipId)).ToList();
                        } catch (InvalidOperationException) {
                            continue;
                        }
                        foreach (IdPartPair added in addedParts) {
                            baseline.Key.DeletePart(added.RelationshipId);
                        }
                    }
                });
                return snapshot;
            }

            internal void Restore() {
                for (int index = _restore.Count - 1; index >= 0; index--) _restore[index]();
            }
        }
    }

    public partial class ExcelDocument {
        internal IReadOnlyList<ExcelMutationDiagnostic> GetMutationDiagnostics(
            int maximumDiagnostics,
            CancellationToken cancellationToken = default) {
            var diagnostics = new List<ExcelMutationDiagnostic>(maximumDiagnostics);
            foreach (var error in ValidateDocument(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (diagnostics.Count >= maximumDiagnostics) break;
                diagnostics.Add(new ExcelMutationDiagnostic(
                    "OPENXML_VALIDATION",
                    ExcelMutationDiagnosticSeverity.Error,
                    error.Description ?? "Open XML validation error.",
                    error.Part?.Uri.ToString()));
            }
            cancellationToken.ThrowIfCancellationRequested();
            return diagnostics;
        }
    }
}
