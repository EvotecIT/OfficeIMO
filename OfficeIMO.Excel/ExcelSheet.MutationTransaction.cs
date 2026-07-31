using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        internal ExcelMutationResult ApplyTransactionalMutation(
            Action<CancellationToken> operation,
            int affectedCells,
            ExcelMutationPlanOptions options,
            CancellationToken cancellationToken) {
            ExcelMutationResult? result = null;
            Batch(_ => {
                cancellationToken.ThrowIfCancellationRequested();
                var snapshot = PackageMutationSnapshot.Capture(_excelDocument.WorkbookPartRoot, options.MaximumSnapshotCharacters);
                try {
                    operation(cancellationToken);
                    cancellationToken.ThrowIfCancellationRequested();
                    WorksheetRoot.Save();
                    MarkRequiresSavePreparation();
                    result = new ExcelMutationResult(
                        affectedCells,
                        _excelDocument.GetMutationDiagnostics(options.MaximumDiagnostics));
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

                AddRoot(workbookPart.Workbook, value => workbookPart.Workbook = value);
                AddRoot(workbookPart.WorkbookStylesPart?.Stylesheet, value => workbookPart.WorkbookStylesPart!.Stylesheet = value);
                AddRoot(workbookPart.SharedStringTablePart?.SharedStringTable, value => workbookPart.SharedStringTablePart!.SharedStringTable = value);
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
                foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts) {
                    AddRoot(worksheetPart.Worksheet, value => worksheetPart.Worksheet = value);
                    AddRoot(worksheetPart.WorksheetCommentsPart?.Comments, value => worksheetPart.WorksheetCommentsPart!.Comments = value);
                    AddRoot(worksheetPart.DrawingsPart?.WorksheetDrawing, value => worksheetPart.DrawingsPart!.WorksheetDrawing = value);
                    foreach (TableDefinitionPart part in worksheetPart.TableDefinitionParts) AddRoot(part.Table, value => part.Table = value);
                    foreach (ChartPart part in worksheetPart.DrawingsPart?.ChartParts ?? Enumerable.Empty<ChartPart>()) AddRoot(part.ChartSpace, value => part.ChartSpace = value);
                    foreach (PivotTablePart part in worksheetPart.PivotTableParts) AddRoot(part.PivotTableDefinition, value => part.PivotTableDefinition = value);
                }
                foreach (PivotTableCacheDefinitionPart part in workbookPart.PivotTableCacheDefinitionParts) {
                    AddRoot(part.PivotCacheDefinition, value => part.PivotCacheDefinition = value);
                    PivotTableCacheRecordsPart? records = part.PivotTableCacheRecordsPart;
                    AddRoot(records?.PivotCacheRecords, value => records!.PivotCacheRecords = value);
                }
                foreach (VmlDrawingPart part in workbookPart.WorksheetParts.SelectMany(sheet => sheet.VmlDrawingParts)) {
                    using Stream source = part.GetStream(FileMode.Open, FileAccess.Read);
                    using var buffer = new MemoryStream();
                    source.CopyTo(buffer);
                    byte[] bytes = buffer.ToArray();
                    characters = checked(characters + bytes.Length);
                    if (characters > maximumCharacters) throw new InvalidOperationException($"Transactional snapshot exceeds MaximumSnapshotCharacters ({maximumCharacters}).");
                    snapshot._restore.Add(() => {
                        using var restore = new MemoryStream(bytes, writable: false);
                        part.FeedData(restore);
                    });
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
        internal IReadOnlyList<ExcelMutationDiagnostic> GetMutationDiagnostics(int maximumDiagnostics) {
            return ValidateDocument(DocumentFormat.OpenXml.FileFormatVersions.Microsoft365)
                .Take(maximumDiagnostics)
                .Select(error => new ExcelMutationDiagnostic(
                    "OPENXML_VALIDATION",
                    ExcelMutationDiagnosticSeverity.Error,
                    error.Description ?? "Open XML validation error.",
                    error.Part?.Uri.ToString()))
                .ToArray();
        }
    }
}
