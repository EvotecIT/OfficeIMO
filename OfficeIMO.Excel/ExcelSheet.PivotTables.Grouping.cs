using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {

        private static List<GeneratedPivotGroupingField> BuildGeneratedPivotGroupingFields(
            IReadOnlyList<string> headers,
            IReadOnlyDictionary<int, ExcelPivotGrouping> groupingMap,
            IReadOnlyList<ExcelPivotCalculatedField> calculatedFields) {
            var fields = new List<GeneratedPivotGroupingField>();
            if (groupingMap.Count == 0) {
                return fields;
            }

            var used = new HashSet<string>(headers, StringComparer.OrdinalIgnoreCase);
            foreach (var calculatedField in calculatedFields) {
                used.Add(calculatedField.Name);
            }

            foreach (var pair in groupingMap.OrderBy(pair => pair.Key)) {
                var grouping = pair.Value;
                if (!grouping.HasGeneratedDateLevels) continue;

                int? parentFieldIndex = null;
                foreach (var level in grouping.GeneratedDateLevels) {
                    string fieldName = EnsureUniqueName($"{headers[pair.Key]} {GetDateGroupFieldSuffix(level)}", used);
                    used.Add(fieldName);
                    var generatedGrouping = ExcelPivotGrouping.Date(fieldName, level, grouping.StartDate, grouping.EndDate);
                    int fieldIndex = headers.Count + fields.Count;
                    fields.Add(new GeneratedPivotGroupingField(pair.Key, fieldIndex, parentFieldIndex, fieldName, level, generatedGrouping));
                    parentFieldIndex = fieldIndex;
                }
            }

            return fields;
        }

        private static IReadOnlyDictionary<int, IReadOnlyList<int>> BuildGeneratedPivotGroupingFieldMap(IReadOnlyList<GeneratedPivotGroupingField> generatedFields) {
            if (generatedFields.Count == 0) {
                return EmptyGeneratedPivotGroupingFieldMap;
            }

            var map = new Dictionary<int, IReadOnlyList<int>>();
            foreach (var group in generatedFields.GroupBy(field => field.SourceIndex)) {
                map[group.Key] = group.Select(field => field.FieldIndex).ToArray();
            }

            return map;
        }

        private static List<int> ExpandGeneratedGroupingFieldIndices(
            IReadOnlyList<int> fieldIndices,
            IReadOnlyDictionary<int, IReadOnlyList<int>> generatedFieldsBySource) {
            var expanded = new List<int>();
            foreach (int index in fieldIndices) {
                if (generatedFieldsBySource.TryGetValue(index, out var generatedFields)) {
                    foreach (int generatedIndex in generatedFields) {
                        if (!expanded.Contains(generatedIndex)) {
                            expanded.Add(generatedIndex);
                        }
                    }
                } else if (!expanded.Contains(index)) {
                    expanded.Add(index);
                }
            }

            return expanded;
        }

        private static (List<int> RowFields, List<int> ColumnFields, List<int> PageFields) ExpandGeneratedGroupingFieldIndices(
            IReadOnlyList<int> rowFieldIndices,
            IReadOnlyList<int> columnFieldIndices,
            IReadOnlyList<int> pageFieldIndices,
            IReadOnlyDictionary<int, IReadOnlyList<int>> generatedFieldsBySource) {
            return (
                ExpandGeneratedGroupingFieldIndices(rowFieldIndices, generatedFieldsBySource),
                ExpandGeneratedGroupingFieldIndices(columnFieldIndices, generatedFieldsBySource),
                ExpandGeneratedGroupingFieldIndices(pageFieldIndices, generatedFieldsBySource));
        }

        private static Dictionary<int, ExcelPivotFieldOptions>? BuildPivotFieldOptionMap(IEnumerable<ExcelPivotFieldOptions>? fieldOptions,
            IDictionary<string, int> headerIndex) {
            if (fieldOptions == null) return null;

            Dictionary<int, ExcelPivotFieldOptions>? map = null;
            foreach (var options in fieldOptions) {
                if (options == null || string.IsNullOrWhiteSpace(options.FieldName)) continue;
                int idx = ResolveFieldIndex(options.FieldName, headerIndex, nameof(fieldOptions));
                map ??= new Dictionary<int, ExcelPivotFieldOptions>();
                map[idx] = options;
            }

            return map;
        }

        private static bool[] BuildPivotSharedItemRequirements(
            int sourceFieldCount,
            IReadOnlyList<int> rowFieldIndices,
            IReadOnlyList<int> columnFieldIndices,
            IReadOnlyList<int> pageFieldIndices,
            IReadOnlyDictionary<int, ExcelPivotGrouping> groupingMap,
            IReadOnlyList<ExcelPivotFilter> pivotFilters,
            IDictionary<string, int> headerIndex,
            IReadOnlyDictionary<int, ExcelPivotFieldOptions>? fieldOptionMap) {
            var requirements = new bool[sourceFieldCount];
            MarkPivotSharedItemRequirements(requirements, rowFieldIndices);
            MarkPivotSharedItemRequirements(requirements, columnFieldIndices);
            MarkPivotSharedItemRequirements(requirements, pageFieldIndices);

            foreach (int fieldIndex in groupingMap.Keys) {
                MarkPivotSharedItemRequirement(requirements, fieldIndex);
            }

            foreach (var filter in pivotFilters) {
                if (filter == null || string.IsNullOrWhiteSpace(filter.FieldName)) {
                    continue;
                }

                MarkPivotSharedItemRequirement(requirements, ResolveFieldIndex(filter.FieldName, headerIndex, nameof(pivotFilters)));
            }

            if (fieldOptionMap != null) {
                foreach (int fieldIndex in fieldOptionMap.Keys) {
                    MarkPivotSharedItemRequirement(requirements, fieldIndex);
                }
            }

            return requirements;
        }

        private static void MarkPivotSharedItemRequirements(bool[] requirements, IReadOnlyList<int> fieldIndices) {
            for (int i = 0; i < fieldIndices.Count; i++) {
                MarkPivotSharedItemRequirement(requirements, fieldIndices[i]);
            }
        }

        private static void MarkPivotSharedItemRequirement(bool[] requirements, int fieldIndex) {
            if ((uint)fieldIndex < (uint)requirements.Length) {
                requirements[fieldIndex] = true;
            }
        }

        private static void ExpandGeneratedGroupingFieldOptions(
            IDictionary<int, ExcelPivotFieldOptions>? fieldOptionMap,
            IReadOnlyDictionary<int, IReadOnlyList<int>> generatedFieldsBySource,
            IReadOnlyList<string> allFields,
            IReadOnlyList<IReadOnlyList<string>> allFieldValueMap) {
            if (fieldOptionMap == null || generatedFieldsBySource.Count == 0) {
                return;
            }

            foreach (var pair in generatedFieldsBySource) {
                if (!fieldOptionMap.TryGetValue(pair.Key, out var sourceOptions)) continue;

                fieldOptionMap[pair.Key] = ClonePivotFieldOptions(sourceOptions, sourceOptions.FieldName);
                foreach (int generatedIndex in pair.Value) {
                    if (generatedIndex < 0 || generatedIndex >= allFields.Count || generatedIndex >= allFieldValueMap.Count) continue;
                    fieldOptionMap[generatedIndex] = ClonePivotFieldOptionsForGeneratedField(
                        sourceOptions,
                        allFields[generatedIndex],
                        allFieldValueMap[generatedIndex]);
                }
            }
        }

        private static IReadOnlyList<string>[] BuildPivotTextValueMap(
            IReadOnlyList<PivotFieldValues> fieldValueMap,
            IReadOnlyList<PivotFieldValues> generatedFieldValueMap,
            int calculatedFieldCount,
            int allFieldCount) {
            var textValueMap = new IReadOnlyList<string>[allFieldCount];
            int index = 0;
            for (int i = 0; i < fieldValueMap.Count; i++) {
                textValueMap[index++] = fieldValueMap[i].TextValues;
            }

            for (int i = 0; i < generatedFieldValueMap.Count; i++) {
                textValueMap[index++] = generatedFieldValueMap[i].TextValues;
            }

            for (int i = 0; i < calculatedFieldCount; i++) {
                textValueMap[index++] = Array.Empty<string>();
            }

            return textValueMap;
        }

        private static ExcelPivotFieldOptions ClonePivotFieldOptionsForGeneratedField(
            ExcelPivotFieldOptions sourceOptions,
            string fieldName,
            IReadOnlyList<string> generatedValues) {
            var valueSet = new HashSet<string>(generatedValues, StringComparer.OrdinalIgnoreCase);
            string[] hiddenItems = sourceOptions.HiddenItems.Where(valueSet.Contains).ToArray();
            string[] visibleItems = sourceOptions.VisibleItems.Where(valueSet.Contains).ToArray();
            string? selectedItem = sourceOptions.SelectedItem != null && valueSet.Contains(sourceOptions.SelectedItem)
                ? sourceOptions.SelectedItem
                : null;

            return ClonePivotFieldOptions(sourceOptions, fieldName, hiddenItems, visibleItems, selectedItem);
        }

        private static ExcelPivotFieldOptions ClonePivotFieldOptions(
            ExcelPivotFieldOptions sourceOptions,
            string fieldName,
            IEnumerable<string>? hiddenItems = null,
            IEnumerable<string>? visibleItems = null,
            string? selectedItem = null) {
            return new ExcelPivotFieldOptions(
                fieldName,
                sourceOptions.SortType,
                sourceOptions.NumberFormatId,
                sourceOptions.NumberFormat,
                sourceOptions.ShowAll,
                sourceOptions.DefaultSubtotal,
                sourceOptions.SubtotalTop,
                sourceOptions.InsertBlankRow,
                sourceOptions.InsertPageBreak,
                sourceOptions.Compact,
                sourceOptions.Outline,
                sourceOptions.ShowDropDowns,
                sourceOptions.MultipleItemSelectionAllowed,
                sourceOptions.IncludeNewItemsInFilter,
                sourceOptions.SubtotalCaption,
                hiddenItems,
                visibleItems,
                selectedItem);
        }

        private static Dictionary<int, ExcelPivotGrouping> BuildPivotGroupingMap(IEnumerable<ExcelPivotGrouping>? groupings,
            IDictionary<string, int> headerIndex,
            int sourceFieldCount) {
            var map = new Dictionary<int, ExcelPivotGrouping>();
            if (groupings == null) return map;

            foreach (var grouping in groupings) {
                if (grouping == null) continue;
                int idx = ResolveFieldIndex(grouping.FieldName, headerIndex, nameof(groupings));
                if (idx >= sourceFieldCount) {
                    throw new ArgumentException($"Pivot grouping field '{grouping.FieldName}' must be a source field, not a calculated field.", nameof(groupings));
                }

                map[idx] = grouping;
            }

            return map;
        }

        private static List<ExcelPivotCalculatedField> NormalizeCalculatedFields(IEnumerable<ExcelPivotCalculatedField>? calculatedFields,
            IReadOnlyList<string> sourceHeaders) {
            var list = new List<ExcelPivotCalculatedField>();
            if (calculatedFields == null) return list;

            var names = new HashSet<string>(sourceHeaders, StringComparer.OrdinalIgnoreCase);
            foreach (var field in calculatedFields) {
                if (field == null) continue;
                if (!names.Add(field.Name)) {
                    throw new ArgumentException($"Pivot calculated field '{field.Name}' duplicates an existing source or calculated field name.", nameof(calculatedFields));
                }

                list.Add(field);
            }

            return list;
        }
    }
}
