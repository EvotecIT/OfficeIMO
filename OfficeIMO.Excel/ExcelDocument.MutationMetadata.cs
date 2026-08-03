using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static void RewriteMutationViewReferences(
            Worksheet? worksheet,
            Func<ExcelReference, ExcelReference?> transform) {
            if (worksheet == null) return;
            foreach (OpenXmlElement view in worksheet.Descendants().Where(element =>
                element is SheetView || element is CustomSheetView)) {
                RewriteMutationSingleReferenceAttribute(view, "topLeftCell", transform, removeOnDeleted: true);
            }
            foreach (Pane pane in worksheet.Descendants<Pane>()) {
                RewriteMutationSingleReferenceAttribute(pane, "topLeftCell", transform, removeOnDeleted: true);
            }
            foreach (Selection selection in worksheet.Descendants<Selection>()) {
                string originalReferences = selection.SequenceOfReferences?.InnerText ?? string.Empty;
                string rewrittenReferences = string.IsNullOrWhiteSpace(originalReferences)
                    ? string.Empty
                    : RewriteReferenceList(originalReferences, transform);
                string? originalActiveCell = selection.ActiveCell?.Value;
                string? rewrittenActiveCell = RewriteMutationSingleReference(originalActiveCell, transform);
                if (!string.IsNullOrWhiteSpace(originalReferences) && rewrittenReferences.Length == 0) {
                    rewrittenReferences = rewrittenActiveCell ?? "A1";
                }
                if (!string.IsNullOrWhiteSpace(originalReferences)) {
                    selection.SequenceOfReferences = new ListValue<StringValue> { InnerText = rewrittenReferences };
                }
                if (!string.IsNullOrWhiteSpace(originalActiveCell)) {
                    selection.ActiveCell = rewrittenActiveCell
                        ?? rewrittenReferences.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()
                        ?? "A1";
                }
                string[] references = rewrittenReferences.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
                if (selection.ActiveCellId != null && references.Length > 0) {
                    uint? matchingIndex = FindContainingReferenceIndex(selection.ActiveCell?.Value, references);
                    if (matchingIndex.HasValue) selection.ActiveCellId = matchingIndex.Value;
                    else if (selection.ActiveCellId.Value >= (uint)references.Length) {
                        selection.ActiveCellId = (uint)references.Length - 1U;
                    }
                }
            }
        }

        private static uint? FindContainingReferenceIndex(string? activeCell, IReadOnlyList<string> references) {
            if (!ExcelReference.TryParse(activeCell, out ExcelReference? active)) return null;
            active!.GetBounds(out int row, out int column, out _, out _);
            for (int index = 0; index < references.Count; index++) {
                if (ExcelReference.TryParse(references[index], out ExcelReference? reference)
                    && reference!.Contains(row, column)) return (uint)index;
            }
            return null;
        }

        private static void RewriteMutationDataConsolidationReferences(
            Worksheet? worksheet,
            string editedSheetName,
            Func<ExcelReference, ExcelReference?> transform) {
            if (worksheet == null) return;
            bool changed = false;
            foreach (DataReference source in worksheet.Descendants<DataReference>().ToList()) {
                if (!string.IsNullOrWhiteSpace(source.Id?.Value)
                    || !string.Equals(source.Sheet?.Value, editedSheetName, StringComparison.OrdinalIgnoreCase)
                    || source.Reference?.Value is not string referenceText
                    || !ExcelReference.TryParse(referenceText, out ExcelReference? reference)) continue;
                ExcelReference? rewritten = transform(reference!);
                if (rewritten == null) {
                    source.Remove();
                    changed = true;
                } else if (!string.Equals(referenceText, rewritten.ToString(), StringComparison.OrdinalIgnoreCase)) {
                    source.Reference = rewritten.ToString();
                    changed = true;
                }
            }
            if (!changed) return;
            foreach (DataReferences references in worksheet.Descendants<DataReferences>().ToList()) {
                uint count = (uint)references.Elements<DataReference>().Count();
                if (count == 0U) references.Remove();
                else references.Count = count;
            }
        }

        private static void RewriteMutationScenarioInputs(
            Worksheet? worksheet,
            Func<ExcelReference, ExcelReference?> transform) {
            if (worksheet == null) return;
            foreach (Scenarios scenarios in worksheet.Descendants<Scenarios>().ToList()) {
                List<Scenario> originalScenarios = scenarios.Elements<Scenario>().ToList();
                var removedIndices = new HashSet<int>();
                for (int index = 0; index < originalScenarios.Count; index++) {
                    Scenario scenario = originalScenarios[index];
                    bool inputsChanged = false;
                    foreach (InputCells input in scenario.Elements<InputCells>().ToList()) {
                        if (input.CellReference?.Value is not string inputReference) continue;
                        string? rewritten = RewriteMutationSingleReference(inputReference, transform);
                        if (rewritten == null) {
                            input.Remove();
                            inputsChanged = true;
                        } else if (!string.Equals(inputReference, rewritten, StringComparison.OrdinalIgnoreCase)) {
                            input.CellReference = rewritten;
                            inputsChanged = true;
                        }
                    }
                    uint inputCount = (uint)scenario.Elements<InputCells>().Count();
                    if (inputCount == 0U) {
                        scenario.Remove();
                        removedIndices.Add(index);
                    } else if (inputsChanged) {
                        scenario.Count = inputCount;
                    }
                }
                int survivingCount = originalScenarios.Count - removedIndices.Count;
                if (survivingCount == 0) {
                    scenarios.Remove();
                    continue;
                }
                if (removedIndices.Count == 0) continue;
                if (scenarios.Current?.Value is uint current) {
                    scenarios.Current = RemapMutationScenarioIndex(current, originalScenarios.Count, removedIndices);
                }
                if (scenarios.Show?.Value is uint shown) {
                    scenarios.Show = RemapMutationScenarioIndex(shown, originalScenarios.Count, removedIndices);
                }
            }
        }

        private static uint RemapMutationScenarioIndex(uint index, int originalCount, ISet<int> removedIndices) {
            int oldIndex = index >= (uint)originalCount ? originalCount - 1 : (int)index;
            int newIndex = 0;
            int lastSurvivingNewIndex = 0;
            for (int candidate = 0; candidate < originalCount; candidate++) {
                if (removedIndices.Contains(candidate)) continue;
                lastSurvivingNewIndex = newIndex;
                if (candidate >= oldIndex) return (uint)newIndex;
                newIndex++;
            }
            return (uint)lastSurvivingNewIndex;
        }

        private static void RewriteMutationCellWatchesAndSmartTags(
            Worksheet? worksheet,
            Func<ExcelReference, ExcelReference?> transform) {
            if (worksheet == null) return;
            const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            CellWatches? watches = worksheet.GetFirstChild<CellWatches>();
            if (watches != null) {
                foreach (CellWatch watch in watches.Elements<CellWatch>().ToList()) {
                    string? original = watch.CellReference?.Value;
                    string? rewritten = RewriteMutationSingleReference(original, transform);
                    if (rewritten == null) {
                        watch.Remove();
                    } else if (!string.Equals(original, rewritten, StringComparison.OrdinalIgnoreCase)) {
                        watch.CellReference = rewritten;
                    }
                }
                if (!watches.Elements<CellWatch>().Any()) watches.Remove();
            }

            var affectedContainers = new HashSet<OpenXmlElement>();
            foreach (OpenXmlElement tag in worksheet.Descendants()
                .Where(element => string.Equals(element.LocalName, "cellSmartTag", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(element.NamespaceUri, SpreadsheetNamespace, StringComparison.Ordinal))
                .ToList()) {
                OpenXmlAttribute? attribute = tag.GetAttributes().FirstOrDefault(candidate =>
                    string.Equals(candidate.LocalName, "r", StringComparison.OrdinalIgnoreCase));
                if (!attribute.HasValue || string.IsNullOrEmpty(attribute.Value.LocalName)) continue;
                string? rewritten = RewriteMutationSingleReference(attribute.Value.Value, transform);
                if (tag.Parent is OpenXmlElement container
                    && string.Equals(
                        container.LocalName,
                        "cellSmartTags",
                        StringComparison.OrdinalIgnoreCase)
                    && string.Equals(container.NamespaceUri, SpreadsheetNamespace, StringComparison.Ordinal)) {
                    affectedContainers.Add(container);
                }
                if (rewritten == null) {
                    tag.Remove();
                } else if (!string.Equals(attribute.Value.Value, rewritten, StringComparison.OrdinalIgnoreCase)) {
                    tag.SetAttribute(new OpenXmlAttribute(
                        attribute.Value.Prefix,
                        attribute.Value.LocalName,
                        attribute.Value.NamespaceUri,
                        rewritten));
                }
            }

            foreach (OpenXmlElement container in affectedContainers) {
                uint count = (uint)container.ChildElements.Count(child =>
                    string.Equals(child.LocalName, "cellSmartTag", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(child.NamespaceUri, SpreadsheetNamespace, StringComparison.Ordinal));
                if (count == 0U) {
                    container.Remove();
                    continue;
                }

                OpenXmlAttribute? countAttribute = container.GetAttributes().FirstOrDefault(candidate =>
                    string.Equals(candidate.LocalName, "count", StringComparison.OrdinalIgnoreCase));
                if (countAttribute.HasValue && !string.IsNullOrEmpty(countAttribute.Value.LocalName)) {
                    container.SetAttribute(new OpenXmlAttribute(
                        countAttribute.Value.Prefix,
                        countAttribute.Value.LocalName,
                        countAttribute.Value.NamespaceUri,
                        count.ToString(CultureInfo.InvariantCulture)));
                }
            }
        }

        private static void RewriteMutationWebPublishItems(
            Workbook workbook,
            string editedSheetName,
            Func<ExcelReference, ExcelReference?> transform) {
            WebPublishItems? items = workbook.GetFirstChild<WebPublishItems>();
            if (items == null) return;
            bool changed = false;
            foreach (WebPublishItem item in items.Elements<WebPublishItem>().ToList()) {
                if (item.SourceType?.Value != WebSourceValues.Range
                    || !string.Equals(item.SourceObject?.Value, editedSheetName, StringComparison.OrdinalIgnoreCase)
                    || item.SourceRef?.Value is not string sourceReference) continue;
                string? rewritten = RewriteMutationSingleReference(sourceReference, transform);
                if (rewritten == null) {
                    item.Remove();
                    changed = true;
                } else if (!string.Equals(sourceReference, rewritten, StringComparison.OrdinalIgnoreCase)) {
                    item.SourceRef = rewritten;
                    changed = true;
                }
            }
            if (!changed) return;
            uint count = (uint)items.Elements<WebPublishItem>().Count();
            if (count == 0U) items.Remove();
            else items.Count = count;
        }

        private static void RewriteMutationSingleReferenceAttribute(
            OpenXmlElement element,
            string localName,
            Func<ExcelReference, ExcelReference?> transform,
            bool removeOnDeleted) {
            OpenXmlAttribute? attribute = element.GetAttributes().FirstOrDefault(candidate =>
                string.Equals(candidate.LocalName, localName, StringComparison.OrdinalIgnoreCase));
            if (!attribute.HasValue || string.IsNullOrEmpty(attribute.Value.LocalName)) return;
            string? rewritten = RewriteMutationSingleReference(attribute.Value.Value, transform);
            if (rewritten == null) {
                if (removeOnDeleted) element.RemoveAttribute(attribute.Value.LocalName, attribute.Value.NamespaceUri);
                return;
            }
            element.SetAttribute(new OpenXmlAttribute(
                attribute.Value.Prefix,
                attribute.Value.LocalName,
                attribute.Value.NamespaceUri,
                rewritten));
        }

        private static string? RewriteMutationSingleReference(
            string? value,
            Func<ExcelReference, ExcelReference?> transform) {
            if (string.IsNullOrWhiteSpace(value)
                || !ExcelReference.TryParse(value, out ExcelReference? reference)) return value;
            return transform(reference!)?.ToString();
        }
    }
}
