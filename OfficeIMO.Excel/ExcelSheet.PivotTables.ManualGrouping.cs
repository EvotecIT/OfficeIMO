using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Adds a derived manual text grouping field to a pivot that uses the source field on one axis.
        /// Named groups contain distinct source item captions; other items remain individual groups.
        /// Optionally hides named or ungrouped items on the derived field.
        /// Materialize or refresh the pivot to populate its displayed values.
        /// </summary>
        public void AddPivotManualGrouping(string pivotTableName, string sourceFieldName, string groupFieldName,
            IReadOnlyDictionary<string, string[]> groups, IReadOnlyCollection<string>? hiddenGroupItems = null) {
            if (string.IsNullOrWhiteSpace(pivotTableName)) throw new ArgumentException("A pivot name is required.", nameof(pivotTableName));
            if (string.IsNullOrWhiteSpace(sourceFieldName)) throw new ArgumentException("A source field is required.", nameof(sourceFieldName));
            if (string.IsNullOrWhiteSpace(groupFieldName)) throw new ArgumentException("A group field name is required.", nameof(groupFieldName));
            if (groups == null || groups.Count == 0) throw new ArgumentException("At least one group is required.", nameof(groups));
            WriteLockWorksheetPreparationOnly(() => {
                PivotTablePart part = _worksheetPart.PivotTableParts.FirstOrDefault(p =>
                    string.Equals(p.PivotTableDefinition?.Name?.Value, pivotTableName, StringComparison.OrdinalIgnoreCase))
                    ?? throw new ArgumentException("The pivot table was not found on this worksheet.", nameof(pivotTableName));
                PivotTableDefinition definition = part.PivotTableDefinition!;
                PivotCacheDefinition cache = part.PivotTableCacheDefinitionPart?.PivotCacheDefinition
                    ?? throw new InvalidOperationException("The pivot cache is missing.");
                if (WorkbookPartRoot.WorksheetParts.SelectMany(sheet => sheet.PivotTableParts)
                    .Count(pivot => ReferenceEquals(pivot.PivotTableCacheDefinitionPart, part.PivotTableCacheDefinitionPart)) != 1)
                    throw new NotSupportedException("Manual grouping requires an unshared pivot cache.");
                CacheField[] fields = cache.CacheFields?.Elements<CacheField>().ToArray() ?? Array.Empty<CacheField>();
                PivotField[] pivotFields = definition.PivotFields?.Elements<PivotField>().ToArray() ?? Array.Empty<PivotField>();
                int source = Array.FindIndex(fields, field => string.Equals(field.Name?.Value, sourceFieldName, StringComparison.OrdinalIgnoreCase));
                if (source < 0 || source >= pivotFields.Length || fields.Length >= 256
                    || fields.Any(field => string.Equals(field.Name?.Value, groupFieldName, StringComparison.OrdinalIgnoreCase))
                    || fields[source].DatabaseField?.Value == false || fields[source].FieldGroup != null)
                    throw new ArgumentException("The source and derived pivot fields must be distinct, ungrouped fields.", nameof(sourceFieldName));
                var row = definition.RowFields?.Elements<Field>().ToList() ?? new List<Field>();
                var column = definition.ColumnFields?.Elements<Field>().ToList() ?? new List<Field>();
                bool onRow = row.Count(field => field.Index?.Value == source) == 1;
                bool onColumn = column.Count(field => field.Index?.Value == source) == 1;
                if (onRow == onColumn)
                    throw new NotSupportedException("Manual grouping requires a source field on one row or column axis.");
                OpenXmlElement[] sourceItems = fields[source].SharedItems?.ChildElements.ToArray() ?? Array.Empty<OpenXmlElement>();
                if (sourceItems.Length == 0 || sourceItems.Length > 100_000 || sourceItems.Any(item => item is not StringItem))
                    throw new NotSupportedException("Manual grouping requires bounded text source items.");
                var sourceValues = new PivotFieldValues(sourceItems.Cast<StringItem>()
                    .Select(item => PivotFieldValue.FromText(item.Val?.Value ?? string.Empty)).ToArray());
                var sourceKeys = new HashSet<PivotFieldValue>(sourceValues.Items);
                if (sourceKeys.Count != sourceValues.Items.Count)
                    throw new NotSupportedException("Manual grouping requires distinct text source items.");
                var assignments = new Dictionary<PivotFieldValue, PivotFieldValue>();
                var labels = new List<PivotFieldValue>();
                var usedLabels = new HashSet<PivotFieldValue>();
                foreach (var entry in groups) {
                    if (string.IsNullOrWhiteSpace(entry.Key) || entry.Value == null || entry.Value.Length < 2)
                        throw new ArgumentException("Each named group needs a label and at least two members.", nameof(groups));
                    var label = PivotFieldValue.FromText(entry.Key);
                    if (!usedLabels.Add(label)) throw new ArgumentException("Group labels must be distinct.", nameof(groups));
                    labels.Add(label);
                    foreach (string member in entry.Value) {
                        if (string.IsNullOrEmpty(member)) throw new ArgumentException("Group members must be nonempty text.", nameof(groups));
                        var key = PivotFieldValue.FromText(member);
                        if (!sourceKeys.Contains(key) || assignments.ContainsKey(key))
                            throw new ArgumentException("Each group member must occur once in the source pivot items.", nameof(groups));
                        assignments.Add(key, label);
                    }
                }
                foreach (var key in sourceValues.Items.Where(key => !assignments.ContainsKey(key))
                    .OrderBy(key => key.Text, StringComparer.OrdinalIgnoreCase)) {
                    if (!usedLabels.Add(key)) throw new ArgumentException("A group label conflicts with an ungrouped item.", nameof(groups));
                    assignments.Add(key, key);
                    labels.Add(key);
                }
                var hiddenLabels = new HashSet<PivotFieldValue>();
                if (hiddenGroupItems != null) {
                    foreach (string hidden in hiddenGroupItems) {
                        if (string.IsNullOrWhiteSpace(hidden) || !hiddenLabels.Add(PivotFieldValue.FromText(hidden)))
                            throw new ArgumentException("Hidden group items must be distinct nonempty labels.", nameof(hiddenGroupItems));
                    }
                }
                if (hiddenLabels.Any(label => !usedLabels.Contains(label)) || hiddenLabels.Count == labels.Count)
                    throw new ArgumentException("Hidden items must name existing groups and leave at least one visible item.", nameof(hiddenGroupItems));
                var groupValues = new PivotFieldValues(labels);
                var groupIndices = new Dictionary<PivotFieldValue, uint>();
                for (uint i = 0; i < labels.Count; i++) groupIndices.Add(labels[(int)i], i);
                var discrete = new DiscreteProperties { Count = (uint)sourceItems.Length };
                foreach (var key in sourceValues.Items) discrete.AppendChild(new FieldItem { Val = groupIndices[assignments[key]] });
                int derived = fields.Length;
                fields[source].FieldGroup = new FieldGroup { ParentId = (uint)derived };
                var derivedField = new CacheField { Name = groupFieldName, DatabaseField = false,
                    FieldGroup = new FieldGroup { Base = (uint)source } };
                derivedField.FieldGroup.Append(discrete, BuildGroupItems(groupValues));
                cache.CacheFields!.Append(derivedField);
                cache.CacheFields.Count = (uint)(derived + 1);
                if (pivotFields[source].Items == null)
                    pivotFields[source].Items = CreateMaterializedFilteredItems(sourceValues, null, false, true);
                pivotFields[source].ShowAll = false;
                var newField = new PivotField { Axis = onRow ? PivotTableAxisValues.AxisRow : PivotTableAxisValues.AxisColumn,
                    ShowAll = false, Compact = false, Outline = false, DefaultSubtotal = true,
                    Items = CreateMaterializedFilteredItems(groupValues, hiddenLabels, false, true) };
                definition.PivotFields!.Append(newField);
                definition.PivotFields.Count = (uint)(derived + 1);
                OpenXmlCompositeElement axis = onRow ? definition.RowFields! : definition.ColumnFields!;
                Field baseAxis = axis.Elements<Field>().First(field => field.Index?.Value == source);
                axis.InsertBefore(new Field { Index = derived }, baseAxis);
                if (onRow) definition.RowFields!.Count = (uint)axis.ChildElements.Count;
                else definition.ColumnFields!.Count = (uint)axis.ChildElements.Count;
            });
        }

        private sealed class PivotManualGrouping {
            internal int SourceField;
            internal OpenXmlElement[] SavedItems = Array.Empty<OpenXmlElement>();
            internal PivotFieldValues Labels = null!;
            internal Dictionary<PivotFieldValue, PivotFieldValue> LabelsBySource = null!;

            internal PivotFieldValue Group(PivotFieldValue source) {
                if (source.Kind != PivotFieldValueKind.Text)
                    throw new NotSupportedException("Manual pivot grouping requires text source keys.");
                return LabelsBySource.TryGetValue(source, out var label) ? label
                    : throw new NotSupportedException("The manual pivot group does not cover a source item.");
            }

            internal void IncludeSourceKeys(PivotFieldValues sourceValues) {
                var labels = Labels.Items.ToList();
                var used = new HashSet<PivotFieldValue>(labels);
                foreach (var source in sourceValues.Items) {
                    if (source.Kind != PivotFieldValueKind.Text)
                        throw new NotSupportedException("Manual pivot grouping requires text source keys.");
                    if (LabelsBySource.ContainsKey(source)) continue;
                    if (!used.Add(source))
                        throw new NotSupportedException("A new source item conflicts with an existing manual group label.");
                    LabelsBySource.Add(source, source);
                    labels.Add(source);
                }
                Labels = new PivotFieldValues(labels);
            }

            internal void RewriteCacheField(CacheField field, PivotFieldValues sourceValues) {
                FieldGroup group = field.FieldGroup ?? throw new NotSupportedException("The manual pivot field group is missing.");
                var index = new Dictionary<PivotFieldValue, int>(Labels.Items.Count);
                for (int item = 0; item < Labels.Items.Count; item++) index.Add(Labels.Items[item], item);
                var discrete = new DiscreteProperties { Count = (uint)sourceValues.Items.Count };
                foreach (var source in sourceValues.Items)
                    discrete.AppendChild(new FieldItem { Val = (uint)index[LabelsBySource[source]] });
                group.RemoveAllChildren<DiscreteProperties>();
                group.RemoveAllChildren<GroupItems>();
                group.Append(discrete);
                group.Append(BuildGroupItems(Labels));
            }
        }

        private static PivotFieldValues OrderManualSourceValues(PivotFieldValues current, CacheField source,
            PivotField pivotField) {
            var saved = source.SharedItems?.ChildElements.ToArray() ?? Array.Empty<OpenXmlElement>();
            var available = new HashSet<PivotFieldValue>(current.Items);
            var ordered = new List<PivotFieldValue>(current.Items.Count);
            foreach (var item in pivotField.Items?.Elements<Item>() ?? Enumerable.Empty<Item>()) {
                if (item.ItemType?.Value == ItemValues.Default) continue;
                if (item.Index == null || item.Index.Value >= saved.Length || saved[item.Index.Value] is not StringItem text)
                    throw new NotSupportedException("The manual group source pivot items are invalid.");
                var key = PivotFieldValue.FromText(text.Val?.Value ?? string.Empty);
                if (available.Remove(key)) ordered.Add(key);
            }
            foreach (var value in current.Items) if (available.Remove(value)) ordered.Add(value);
            return new PivotFieldValues(ordered);
        }

        private static bool IsDerivedManualGroup(CacheField field) => field.DatabaseField?.Value == false
            && field.FieldGroup?.GetFirstChild<DiscreteProperties>() != null;

        private static PivotManualGrouping ReadPivotManualGrouping(CacheField[] fields, PivotField[] pivotFields,
            int index, int sourceFieldCount) {
            CacheField field = fields[index];
            FieldGroup? group = field.FieldGroup;
            DiscreteProperties? discrete = group?.GetFirstChild<DiscreteProperties>();
            if (!IsDerivedManualGroup(field) || field.Formula != null || group?.Base == null
                || group.Base.Value >= sourceFieldCount || group.ParentId != null
                || group.GetFirstChild<RangeProperties>() != null || discrete == null)
                throw new NotSupportedException("This derived manual pivot group is not qualified for materialization or lookup.");
            int sourceField = (int)group.Base.Value;
            CacheField source = fields[sourceField];
            if (source.DatabaseField?.Value == false || source.Formula != null
                || source.FieldGroup?.ParentId?.Value != index
                || source.FieldGroup?.GetFirstChild<RangeProperties>() != null)
                throw new NotSupportedException("The manual pivot group must reference its source field.");
            OpenXmlElement[] sourceItems = source.SharedItems?.ChildElements.ToArray() ?? Array.Empty<OpenXmlElement>();
            GroupItems? groupItems = group!.GetFirstChild<GroupItems>();
            OpenXmlElement[] savedItems = groupItems?.ChildElements.ToArray()
                ?? Array.Empty<OpenXmlElement>();
            FieldItem[] mapping = discrete.Elements<FieldItem>().ToArray();
            if (sourceItems.Length == 0 || sourceItems.Length > 100_000 || savedItems.Length == 0
                || savedItems.Length > 100_000 || mapping.Length != sourceItems.Length
                || discrete.Count?.Value != mapping.Length || groupItems?.Count?.Value != savedItems.Length
                || mapping.Length != discrete.ChildElements.Count || sourceItems.Any(item => item is not StringItem)
                || savedItems.Any(item => item is not StringItem))
                throw new NotSupportedException("The manual pivot group requires bounded text items and a complete mapping.");
            var captions = new PivotFieldValue?[savedItems.Length];
            var ordered = new List<PivotFieldValue>(savedItems.Length);
            Item[] visibleItems = pivotFields[index].Items?.Elements<Item>().ToArray() ?? Array.Empty<Item>();
            foreach (Item item in visibleItems) {
                if (item.ItemType?.Value == ItemValues.Default) continue;
                if (item.ItemType?.Value is ItemValues type && type != ItemValues.Data
                    || item.Index == null || item.Index.Value >= savedItems.Length
                    || captions[item.Index.Value] != null)
                    throw new NotSupportedException("The manual pivot field items do not match its group cache.");
                string caption = PivotLookupAttribute(item, "n");
                if (caption.Length == 0) caption = ((StringItem)savedItems[item.Index.Value]).Val?.Value ?? string.Empty;
                if (caption.Length == 0)
                    throw new NotSupportedException("A manual pivot group label is empty.");
                var label = PivotFieldValue.FromText(caption);
                captions[item.Index.Value] = label;
                ordered.Add(label);
            }
            if (ordered.Count != savedItems.Length || new HashSet<PivotFieldValue>(ordered).Count != ordered.Count)
                throw new NotSupportedException("The manual pivot group labels are missing or ambiguous.");
            var bySource = new Dictionary<PivotFieldValue, PivotFieldValue>(sourceItems.Length);
            for (int item = 0; item < sourceItems.Length; item++) {
                uint? groupIndex = mapping[item].Val?.Value;
                if (groupIndex == null || groupIndex.Value >= captions.Length)
                    throw new NotSupportedException("The manual pivot mapping references an unknown group item.");
                var key = PivotFieldValue.FromText(((StringItem)sourceItems[item]).Val?.Value ?? string.Empty);
                if (bySource.ContainsKey(key))
                    throw new NotSupportedException("The manual pivot source items are ambiguous.");
                bySource.Add(key, captions[groupIndex.Value]!);
            }
            return new PivotManualGrouping { SourceField = sourceField, SavedItems = savedItems,
                Labels = new PivotFieldValues(ordered), LabelsBySource = bySource };
        }
    }
}
