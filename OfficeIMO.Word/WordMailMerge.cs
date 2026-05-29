using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.CustomXmlDataProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

namespace OfficeIMO.Word {
    /// <summary>
    /// Describes a mail-merge template validation issue.
    /// </summary>
    public enum WordMailMergeTemplateIssueKind {
        /// <summary>A MERGEFIELD was found without a supplied value.</summary>
        MissingMergeFieldValue,

        /// <summary>A conditional template block was found without a supplied condition.</summary>
        MissingConditionalValue,

        /// <summary>A conditional block start marker did not have a matching end marker.</summary>
        UnmatchedConditionalStart,

        /// <summary>A conditional block end marker did not have a matching start marker.</summary>
        UnmatchedConditionalEnd,

        /// <summary>A conditional block end marker closed a different block name than the current start marker.</summary>
        MismatchedConditionalEnd,

        /// <summary>A repeated template block was found without supplied rows.</summary>
        MissingRepeatingBlockData,

        /// <summary>A repeated block start marker did not have a matching end marker.</summary>
        UnmatchedRepeatingBlockStart,

        /// <summary>A repeated block end marker did not have a matching start marker.</summary>
        UnmatchedRepeatingBlockEnd,

        /// <summary>A repeated block end marker closed a different block name than the current start marker.</summary>
        MismatchedRepeatingBlockEnd
    }

    /// <summary>
    /// Represents one mail-merge template validation issue.
    /// </summary>
    public sealed class WordMailMergeTemplateIssue {
        internal WordMailMergeTemplateIssue(WordMailMergeTemplateIssueKind kind, string name, string message) {
            Kind = kind;
            Name = name;
            Message = message;
        }

        /// <summary>Issue category.</summary>
        public WordMailMergeTemplateIssueKind Kind { get; }

        /// <summary>Field or conditional block name related to the issue.</summary>
        public string Name { get; }

        /// <summary>Human-readable issue text.</summary>
        public string Message { get; }
    }

    /// <summary>
    /// Describes the merge fields, conditional blocks, and validation issues found in a Word mail-merge template.
    /// </summary>
    public sealed class WordMailMergeTemplateInspection {
        internal WordMailMergeTemplateInspection(IEnumerable<string> mergeFieldNames, IEnumerable<string> conditionalBlockNames, IEnumerable<string> repeatingBlockNames, IEnumerable<WordMailMergeTemplateIssue> issues) {
            MergeFieldNames = mergeFieldNames.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToList();
            ConditionalBlockNames = conditionalBlockNames.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToList();
            RepeatingBlockNames = repeatingBlockNames.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToList();
            Issues = issues.ToList();
        }

        /// <summary>Unique MERGEFIELD names found in the template.</summary>
        public IReadOnlyList<string> MergeFieldNames { get; }

        /// <summary>Unique conditional block names found in the template.</summary>
        public IReadOnlyList<string> ConditionalBlockNames { get; }

        /// <summary>Unique repeated block names found in the template.</summary>
        public IReadOnlyList<string> RepeatingBlockNames { get; }

        /// <summary>Validation issues found during inspection.</summary>
        public IReadOnlyList<WordMailMergeTemplateIssue> Issues { get; }

        /// <summary>True when the template has no validation issues.</summary>
        public bool IsValid => Issues.Count == 0;

        /// <summary>
        /// Throws when validation issues were found, otherwise returns this inspection.
        /// </summary>
        public WordMailMergeTemplateInspection EnsureValid() {
            if (!IsValid) {
                throw new InvalidOperationException(string.Join(Environment.NewLine, Issues.Select(issue => issue.Message)));
            }

            return this;
        }
    }

    /// <summary>
    /// Summarizes a content-control data-binding fill or refresh operation.
    /// </summary>
    public sealed class WordContentControlDataBindingResult {
        internal WordContentControlDataBindingResult(int bindingCount, int updatedContentControls, int updatedCustomXmlNodes, IReadOnlyList<string> missingValueKeys) {
            BindingCount = bindingCount;
            UpdatedContentControls = updatedContentControls;
            UpdatedCustomXmlNodes = updatedCustomXmlNodes;
            MissingValueKeys = missingValueKeys;
        }

        /// <summary>Number of bound content controls found in the document.</summary>
        public int BindingCount { get; }

        /// <summary>Number of bound content controls whose visible text was updated.</summary>
        public int UpdatedContentControls { get; }

        /// <summary>Number of backing Custom XML nodes updated from supplied values.</summary>
        public int UpdatedCustomXmlNodes { get; }

        /// <summary>Binding keys that could not be resolved from supplied values or backing Custom XML.</summary>
        public IReadOnlyList<string> MissingValueKeys { get; }

        /// <summary>True when one or more bound controls had no resolvable value.</summary>
        public bool HasMissingValues => MissingValueKeys.Count > 0;
    }

    /// <summary>
    /// Represents one grouped table-row mail-merge data set.
    /// </summary>
    public sealed class WordMailMergeTableRowGroup {
        /// <summary>
        /// Creates a grouped table-row data set.
        /// </summary>
        /// <param name="values">Values applied to the group template row.</param>
        /// <param name="rows">Values applied to repeated detail rows inside the group.</param>
        public WordMailMergeTableRowGroup(IDictionary<string, string> values, IEnumerable<IDictionary<string, string>> rows) {
            Values = values ?? throw new ArgumentNullException(nameof(values));
            Rows = rows ?? throw new ArgumentNullException(nameof(rows));
        }

        /// <summary>Values applied to the group template row.</summary>
        public IDictionary<string, string> Values { get; }

        /// <summary>Values applied to repeated detail rows inside the group.</summary>
        public IEnumerable<IDictionary<string, string>> Rows { get; }
    }

    /// <summary>
    /// Summarizes a grouped table-row mail-merge operation.
    /// </summary>
    public sealed class WordMailMergeTableRowGroupResult {
        internal WordMailMergeTableRowGroupResult(int groupCount, int detailRowCount) {
            GroupCount = groupCount;
            DetailRowCount = detailRowCount;
        }

        /// <summary>Number of generated group rows.</summary>
        public int GroupCount { get; }

        /// <summary>Number of generated detail rows across all groups.</summary>
        public int DetailRowCount { get; }

        /// <summary>Total number of generated rows.</summary>
        public int TotalRowCount => GroupCount + DetailRowCount;
    }

    /// <summary>
    /// Represents one repeated block row with optional nested repeated regions.
    /// </summary>
    public sealed class WordMailMergeBlockData {
        /// <summary>
        /// Creates a repeated block row.
        /// </summary>
        /// <param name="values">Values applied to merge fields inside this block row.</param>
        public WordMailMergeBlockData(IDictionary<string, string> values)
            : this(values, new Dictionary<string, IEnumerable<WordMailMergeBlockData>>(StringComparer.OrdinalIgnoreCase)) {
        }

        /// <summary>
        /// Creates a repeated block row with nested repeated regions.
        /// </summary>
        /// <param name="values">Values applied to merge fields inside this block row.</param>
        /// <param name="regions">Nested repeated regions available inside this block row.</param>
        public WordMailMergeBlockData(IDictionary<string, string> values, IDictionary<string, IEnumerable<WordMailMergeBlockData>> regions) {
            Values = values ?? throw new ArgumentNullException(nameof(values));
            Regions = regions ?? throw new ArgumentNullException(nameof(regions));
        }

        /// <summary>Values applied to merge fields inside this block row.</summary>
        public IDictionary<string, string> Values { get; }

        /// <summary>Nested repeated regions available inside this block row.</summary>
        public IDictionary<string, IEnumerable<WordMailMergeBlockData>> Regions { get; }
    }

    /// <summary>
    /// Provides basic mail merge capabilities by replacing <c>MERGEFIELD</c> fields with supplied values.
    /// </summary>
    public static class WordMailMerge {
        private static readonly Regex ConditionalBlockMarkerRegex = new Regex(
            @"^\s*\{\{\s*(?<kind>[#/])\s*(?<name>[A-Za-z0-9_.-]+)\s*\}\}\s*$",
            RegexOptions.Compiled,
            TimeSpan.FromMilliseconds(100));

        private static readonly Regex RepeatingBlockMarkerRegex = new Regex(
            @"^\s*\{\{\s*(?<kind>#each|/each)\s+(?<name>[A-Za-z0-9_.-]+)\s*\}\}\s*$",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            TimeSpan.FromMilliseconds(100));

        /// <summary>
        /// Replaces all MERGEFIELD fields in the given document with provided values.
        /// </summary>
        /// <param name="document">Document to update.</param>
        /// <param name="values">Dictionary with field names and values.</param>
        /// <param name="removeFields">Determines whether the field codes are removed after replacement.</param>
        public static void Execute(WordDocument document, IDictionary<string, string> values, bool removeFields = true) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (values == null) throw new ArgumentNullException(nameof(values));

            foreach (var root in EnumerateTemplateRoots(document)) {
                ReplaceMergeFields(root, values, removeFields);
            }
        }

        /// <summary>
        /// Creates one merged document per value set from a template file.
        /// </summary>
        /// <param name="templatePath">Path to the source template document.</param>
        /// <param name="records">Value sets used to generate output documents.</param>
        /// <param name="outputPathFactory">Function that returns an output path for each zero-based record index and value set.</param>
        /// <param name="removeFields">Determines whether field codes are removed after replacement.</param>
        /// <returns>Output document paths in generation order.</returns>
        public static IReadOnlyList<string> ExecuteBatch(string templatePath, IEnumerable<IDictionary<string, string>> records, Func<int, IDictionary<string, string>, string> outputPathFactory, bool removeFields = true) {
            if (string.IsNullOrWhiteSpace(templatePath)) throw new ArgumentException("Template path cannot be empty.", nameof(templatePath));
            if (!File.Exists(templatePath)) throw new FileNotFoundException("Template document was not found.", templatePath);
            if (records == null) throw new ArgumentNullException(nameof(records));
            if (outputPathFactory == null) throw new ArgumentNullException(nameof(outputPathFactory));

            var recordList = records.ToList();
            var outputPaths = new List<string>(recordList.Count);
            for (int index = 0; index < recordList.Count; index++) {
                IDictionary<string, string> values = recordList[index] ?? throw new ArgumentException("Records cannot contain null value dictionaries.", nameof(records));
                string outputPath = outputPathFactory(index, values);
                if (string.IsNullOrWhiteSpace(outputPath)) {
                    throw new InvalidOperationException($"Output path for record {index} cannot be empty.");
                }

                string? directory = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrWhiteSpace(directory)) {
                    Directory.CreateDirectory(directory!);
                }

                using (WordDocument document = WordDocument.Load(templatePath)) {
                    Execute(document, values, removeFields);
                    document.Save(outputPath, false);
                }

                outputPaths.Add(outputPath);
            }

            return outputPaths;
        }

        /// <summary>
        /// Repeats a table row template once for every supplied value set and applies MERGEFIELD replacements in each generated row.
        /// </summary>
        /// <param name="table">Table containing the row template.</param>
        /// <param name="templateRowIndex">Zero-based row index to clone and bind.</param>
        /// <param name="rows">Value sets used to generate rows.</param>
        /// <param name="removeFields">Determines whether field codes are removed after replacement.</param>
        /// <returns>The number of generated rows.</returns>
        public static int ExecuteTableRows(WordTable table, int templateRowIndex, IEnumerable<IDictionary<string, string>> rows, bool removeFields = true) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (rows == null) throw new ArgumentNullException(nameof(rows));

            var tableRows = table.Rows;
            if (templateRowIndex < 0 || templateRowIndex >= tableRows.Count) {
                throw new ArgumentOutOfRangeException(nameof(templateRowIndex), "Template row index is outside the table row range.");
            }

            var rowValues = rows.ToList();
            TableRow templateRow = tableRows[templateRowIndex]._tableRow;

            foreach (var values in rowValues) {
                if (values == null) throw new ArgumentException("Rows cannot contain null value dictionaries.", nameof(rows));

                var clonedRow = (TableRow)templateRow.CloneNode(true);
                ReplaceMergeFields(clonedRow, values, removeFields);
                templateRow.InsertBeforeSelf(clonedRow);
            }

            templateRow.Remove();
            return rowValues.Count;
        }

        /// <summary>
        /// Repeats a group template row and its detail row template for each supplied grouped data set.
        /// </summary>
        /// <param name="table">Table containing the group and detail template rows.</param>
        /// <param name="groupTemplateRowIndex">Zero-based row index of the group/header row template.</param>
        /// <param name="detailTemplateRowIndex">Zero-based row index of the detail row template.</param>
        /// <param name="groups">Grouped value sets used to generate group and detail rows.</param>
        /// <param name="removeFields">Determines whether the field codes are removed after replacement.</param>
        /// <returns>Generated group and detail row counts.</returns>
        public static WordMailMergeTableRowGroupResult ExecuteTableRowGroups(WordTable table, int groupTemplateRowIndex, int detailTemplateRowIndex, IEnumerable<WordMailMergeTableRowGroup> groups, bool removeFields = true) {
            if (table == null) throw new ArgumentNullException(nameof(table));
            if (groups == null) throw new ArgumentNullException(nameof(groups));

            var tableRows = table.Rows;
            if (groupTemplateRowIndex < 0 || groupTemplateRowIndex >= tableRows.Count) {
                throw new ArgumentOutOfRangeException(nameof(groupTemplateRowIndex), "Group template row index is outside the table row range.");
            }

            if (detailTemplateRowIndex < 0 || detailTemplateRowIndex >= tableRows.Count) {
                throw new ArgumentOutOfRangeException(nameof(detailTemplateRowIndex), "Detail template row index is outside the table row range.");
            }

            if (groupTemplateRowIndex == detailTemplateRowIndex) {
                throw new ArgumentException("Group and detail template rows must be different rows.", nameof(detailTemplateRowIndex));
            }

            var groupList = groups.ToList();
            TableRow groupTemplateRow = tableRows[groupTemplateRowIndex]._tableRow;
            TableRow detailTemplateRow = tableRows[detailTemplateRowIndex]._tableRow;
            int detailRowCount = 0;

            foreach (var group in groupList) {
                if (group == null) throw new ArgumentException("Groups cannot contain null items.", nameof(groups));

                var clonedGroupRow = (TableRow)groupTemplateRow.CloneNode(true);
                ReplaceMergeFields(clonedGroupRow, group.Values, removeFields);
                groupTemplateRow.InsertBeforeSelf(clonedGroupRow);

                foreach (var rowValues in group.Rows) {
                    if (rowValues == null) throw new ArgumentException("Group rows cannot contain null value dictionaries.", nameof(groups));

                    var clonedDetailRow = (TableRow)detailTemplateRow.CloneNode(true);
                    ReplaceMergeFields(clonedDetailRow, rowValues, removeFields);
                    groupTemplateRow.InsertBeforeSelf(clonedDetailRow);
                    detailRowCount++;
                }
            }

            detailTemplateRow.Remove();
            groupTemplateRow.Remove();
            return new WordMailMergeTableRowGroupResult(groupList.Count, detailRowCount);
        }

        /// <summary>
        /// Repeats block-level template regions delimited by marker paragraphs such as <c>{{#each Items}}</c> and <c>{{/each Items}}</c>.
        /// </summary>
        /// <param name="document">Document containing repeated block marker paragraphs.</param>
        /// <param name="regions">Repeated region rows keyed by marker name.</param>
        /// <param name="removeFields">Determines whether field codes are removed after replacement.</param>
        /// <returns>The number of generated region instances.</returns>
        public static int ExecuteRepeatingBlocks(WordDocument document, IDictionary<string, IEnumerable<IDictionary<string, string>>> regions, bool removeFields = true) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (regions == null) throw new ArgumentNullException(nameof(regions));

            return ExecuteRepeatingBlockRegions(document, ConvertFlatRegions(regions), removeFields);
        }

        /// <summary>
        /// Repeats block-level template regions with nested data, using marker paragraphs such as <c>{{#each Items}}</c> and <c>{{/each Items}}</c>.
        /// </summary>
        /// <param name="document">Document containing repeated block marker paragraphs.</param>
        /// <param name="regions">Repeated region rows keyed by marker name.</param>
        /// <param name="removeFields">Determines whether field codes are removed after replacement.</param>
        /// <returns>The number of generated region instances, including nested instances.</returns>
        public static int ExecuteRepeatingBlockRegions(WordDocument document, IDictionary<string, IEnumerable<WordMailMergeBlockData>> regions, bool removeFields = true) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (regions == null) throw new ArgumentNullException(nameof(regions));

            int generated = 0;
            foreach (var root in EnumerateTemplateRoots(document)) {
                generated += ExecuteRepeatingBlocks(root, regions, removeFields);
            }

            return generated;
        }

        /// <summary>
        /// Includes or removes conditional template blocks delimited by marker paragraphs such as <c>{{#ShowDiscount}}</c> and <c>{{/ShowDiscount}}</c>.
        /// </summary>
        /// <param name="document">Document containing conditional marker paragraphs.</param>
        /// <param name="conditions">Condition values keyed by marker name.</param>
        /// <param name="removeMarkers">When true, matched marker paragraphs are removed from included blocks.</param>
        /// <returns>The number of conditional blocks processed.</returns>
        public static int ExecuteConditionalBlocks(WordDocument document, IDictionary<string, bool> conditions, bool removeMarkers = true) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (conditions == null) throw new ArgumentNullException(nameof(conditions));

            int processed = 0;
            foreach (var root in EnumerateTemplateRoots(document)) {
                processed += ExecuteConditionalBlocks(root, conditions, removeMarkers);
            }

            return processed;
        }

        /// <summary>
        /// Inspects a Word mail-merge template and optionally validates that merge fields and conditional blocks have supplied values.
        /// </summary>
        /// <param name="document">Document to inspect.</param>
        /// <param name="mergeFieldNames">Optional supplied MERGEFIELD names. When provided, missing fields are reported as issues.</param>
        /// <param name="conditionNames">Optional supplied conditional block names. When provided, missing conditions are reported as issues.</param>
        /// <param name="repeatingBlockNames">Optional supplied repeated block names. When provided, missing repeated block rows are reported as issues.</param>
        public static WordMailMergeTemplateInspection InspectTemplate(WordDocument document, IEnumerable<string>? mergeFieldNames = null, IEnumerable<string>? conditionNames = null, IEnumerable<string>? repeatingBlockNames = null) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            var issues = new List<WordMailMergeTemplateIssue>();
            var mergeFields = new List<string>();
            var conditionalNames = new List<string>();
            var repeatingNames = new List<string>();

            foreach (var root in EnumerateTemplateRoots(document)) {
                mergeFields.AddRange(EnumerateMergeFieldNames(root));
                var conditionalInspection = InspectConditionalBlocks(root);
                conditionalNames.AddRange(conditionalInspection.Names);
                issues.AddRange(conditionalInspection.Issues);
                var repeatingInspection = InspectRepeatingBlocks(root);
                repeatingNames.AddRange(repeatingInspection.Names);
                issues.AddRange(repeatingInspection.Issues);
            }

            var suppliedMergeFields = mergeFieldNames == null
                ? null
                : new HashSet<string>(mergeFieldNames.Where(name => !string.IsNullOrWhiteSpace(name)), StringComparer.OrdinalIgnoreCase);
            if (suppliedMergeFields != null) {
                foreach (string fieldName in mergeFields.Distinct(StringComparer.OrdinalIgnoreCase)) {
                    if (!suppliedMergeFields.Contains(fieldName)) {
                        issues.Add(new WordMailMergeTemplateIssue(
                            WordMailMergeTemplateIssueKind.MissingMergeFieldValue,
                            fieldName,
                            $"Merge field '{fieldName}' was not supplied."));
                    }
                }
            }

            var suppliedConditions = conditionNames == null
                ? null
                : new HashSet<string>(conditionNames.Where(name => !string.IsNullOrWhiteSpace(name)), StringComparer.OrdinalIgnoreCase);
            if (suppliedConditions != null) {
                foreach (string conditionName in conditionalNames.Distinct(StringComparer.OrdinalIgnoreCase)) {
                    if (!suppliedConditions.Contains(conditionName)) {
                        issues.Add(new WordMailMergeTemplateIssue(
                            WordMailMergeTemplateIssueKind.MissingConditionalValue,
                            conditionName,
                            $"Conditional block '{conditionName}' was not supplied."));
                    }
                }
            }

            var suppliedRepeatingBlocks = repeatingBlockNames == null
                ? null
                : new HashSet<string>(repeatingBlockNames.Where(name => !string.IsNullOrWhiteSpace(name)), StringComparer.OrdinalIgnoreCase);
            if (suppliedRepeatingBlocks != null) {
                foreach (string repeatingName in repeatingNames.Distinct(StringComparer.OrdinalIgnoreCase)) {
                    if (!suppliedRepeatingBlocks.Contains(repeatingName)) {
                        issues.Add(new WordMailMergeTemplateIssue(
                            WordMailMergeTemplateIssueKind.MissingRepeatingBlockData,
                            repeatingName,
                            $"Repeating block '{repeatingName}' was not supplied."));
                    }
                }
            }

            return new WordMailMergeTemplateInspection(mergeFields, conditionalNames, repeatingNames, issues);
        }

        /// <summary>
        /// Validates a Word mail-merge template against supplied merge-field and conditional-block names.
        /// </summary>
        /// <param name="document">Document to validate.</param>
        /// <param name="mergeFieldNames">Supplied MERGEFIELD names.</param>
        /// <param name="conditionNames">Supplied conditional block names.</param>
        public static WordMailMergeTemplateInspection ValidateTemplate(WordDocument document, IEnumerable<string> mergeFieldNames, IEnumerable<string> conditionNames) {
            if (mergeFieldNames == null) throw new ArgumentNullException(nameof(mergeFieldNames));
            if (conditionNames == null) throw new ArgumentNullException(nameof(conditionNames));

            return InspectTemplate(document, mergeFieldNames, conditionNames);
        }

        /// <summary>
        /// Validates a Word mail-merge template against supplied merge-field, conditional-block, and repeated-block names.
        /// </summary>
        /// <param name="document">Document to validate.</param>
        /// <param name="mergeFieldNames">Supplied MERGEFIELD names.</param>
        /// <param name="conditionNames">Supplied conditional block names.</param>
        /// <param name="repeatingBlockNames">Supplied repeated block names.</param>
        public static WordMailMergeTemplateInspection ValidateTemplate(WordDocument document, IEnumerable<string> mergeFieldNames, IEnumerable<string> conditionNames, IEnumerable<string> repeatingBlockNames) {
            if (mergeFieldNames == null) throw new ArgumentNullException(nameof(mergeFieldNames));
            if (conditionNames == null) throw new ArgumentNullException(nameof(conditionNames));
            if (repeatingBlockNames == null) throw new ArgumentNullException(nameof(repeatingBlockNames));

            return InspectTemplate(document, mergeFieldNames, conditionNames, repeatingBlockNames);
        }

        /// <summary>
        /// Updates visible text in bound content controls from their backing Custom XML values.
        /// </summary>
        /// <param name="document">Document containing bound content controls.</param>
        public static WordContentControlDataBindingResult RefreshContentControlDataBindings(WordDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            return ExecuteContentControlDataBindingsCore(document, values: null, updateCustomXml: false);
        }

        /// <summary>
        /// Fills bound content controls from supplied values and, when possible, updates the backing Custom XML node.
        /// </summary>
        /// <param name="document">Document containing bound content controls.</param>
        /// <param name="values">Values keyed by content-control alias, tag, XPath, or <c>storeItemId|XPath</c>.</param>
        /// <param name="updateCustomXml">When true, matching backing Custom XML nodes are updated with supplied values.</param>
        public static WordContentControlDataBindingResult ExecuteContentControlDataBindings(WordDocument document, IDictionary<string, string> values, bool updateCustomXml = true) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (values == null) throw new ArgumentNullException(nameof(values));

            return ExecuteContentControlDataBindingsCore(document, values, updateCustomXml);
        }

        private static void ReplaceMergeFields(OpenXmlElement root, IDictionary<string, string> values, bool removeFields) {
            ReplaceSimpleMergeFields(root, values, removeFields);
            ReplaceComplexMergeFields(root, values, removeFields);
        }

        private static WordContentControlDataBindingResult ExecuteContentControlDataBindingsCore(WordDocument document, IDictionary<string, string>? values, bool updateCustomXml) {
            MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart
                ?? throw new InvalidOperationException("MainDocumentPart is missing.");
            var customXmlCache = new Dictionary<CustomXmlPart, XDocument>();
            var dirtyCustomXmlParts = new HashSet<CustomXmlPart>();
            var missingKeys = new List<string>();
            int bindingCount = 0;
            int updatedContentControls = 0;
            int updatedCustomXmlNodes = 0;

            foreach (var boundControl in EnumerateBoundContentControls(mainPart)) {
                bindingCount++;

                string? value = null;
                bool hasSuppliedValue = values != null && TryGetSuppliedBindingValue(values, boundControl, out value);
                if (!hasSuppliedValue && !TryGetCustomXmlBindingValue(mainPart, boundControl.Binding, customXmlCache, out value)) {
                    missingKeys.Add(GetBindingDisplayKey(boundControl));
                    continue;
                }

                if (UpdateContentControlText(boundControl.ContentControl, value!)) {
                    updatedContentControls++;
                }

                if (updateCustomXml && hasSuppliedValue
                    && TrySelectCustomXmlBindingElement(mainPart, boundControl.Binding, customXmlCache, out CustomXmlPart? customXmlPart, out XElement? element)) {
                    if (!string.Equals(element.Value, value, StringComparison.Ordinal)) {
                        element.Value = value!;
                        updatedCustomXmlNodes++;
                        dirtyCustomXmlParts.Add(customXmlPart);
                    }
                }
            }

            foreach (CustomXmlPart part in dirtyCustomXmlParts) {
                SaveCustomXmlPart(part, customXmlCache[part]);
            }

            return new WordContentControlDataBindingResult(
                bindingCount,
                updatedContentControls,
                updatedCustomXmlNodes,
                missingKeys.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(key => key, StringComparer.OrdinalIgnoreCase).ToList());
        }

        private static int ExecuteRepeatingBlocks(OpenXmlCompositeElement container, IDictionary<string, IEnumerable<WordMailMergeBlockData>> regions, bool removeFields) {
            var elements = container.ChildElements.ToList();
            var stack = new List<RepeatingBlockStart>();
            var blocks = new List<RepeatingBlockRange>();

            for (int index = 0; index < elements.Count; index++) {
                if (elements[index] is not Paragraph paragraph || !TryGetRepeatingBlockMarker(paragraph, out string? markerName, out bool isStart)) {
                    continue;
                }

                if (isStart) {
                    stack.Add(new RepeatingBlockStart(markerName!, index));
                    continue;
                }

                if (stack.Count == 0) {
                    throw new InvalidOperationException($"Repeating block end marker '{{{{/each {markerName}}}}}' does not have a matching start marker.");
                }

                var start = stack[stack.Count - 1];
                stack.RemoveAt(stack.Count - 1);
                if (!string.Equals(start.Name, markerName, StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidOperationException($"Repeating block end marker '{{{{/each {markerName}}}}}' does not match start marker '{{{{#each {start.Name}}}}}'.");
                }

                if (stack.Count == 0) {
                    blocks.Add(new RepeatingBlockRange(start.Name, start.Index, index));
                }
            }

            if (stack.Count > 0) {
                var start = stack[stack.Count - 1];
                throw new InvalidOperationException($"Repeating block start marker '{{{{#each {start.Name}}}}}' does not have a matching end marker.");
            }

            int generated = 0;
            foreach (var block in blocks.OrderByDescending(item => item.StartIndex)) {
                if (!TryGetRegionRows(regions, block.Name, out IEnumerable<WordMailMergeBlockData>? rows)) {
                    throw new InvalidOperationException($"Repeating block '{block.Name}' was not supplied.");
                }

                var templateElements = elements
                    .Skip(block.StartIndex + 1)
                    .Take(block.EndIndex - block.StartIndex - 1)
                    .ToList();
                var rowValues = rows!.ToList();
                OpenXmlElement insertionPoint = elements[block.StartIndex];

                foreach (var row in rowValues) {
                    if (row == null) throw new ArgumentException("Repeating block rows cannot contain null items.", nameof(regions));

                    var workingContainer = new SdtContentBlock();
                    foreach (OpenXmlElement templateElement in templateElements) {
                        var clonedElement = templateElement.CloneNode(true);
                        workingContainer.Append(clonedElement);
                    }

                    generated += ExecuteRepeatingBlocks(workingContainer, row.Regions, removeFields);
                    ReplaceMergeFields(workingContainer, row.Values, removeFields);

                    foreach (OpenXmlElement clonedElement in workingContainer.ChildElements.ToList()) {
                        clonedElement.Remove();
                        insertionPoint.InsertBeforeSelf(clonedElement);
                    }

                    generated++;
                }

                for (int index = block.StartIndex; index <= block.EndIndex; index++) {
                    RemoveIfAttached(elements[index]);
                }
            }

            foreach (var child in container.ChildElements.OfType<OpenXmlCompositeElement>().ToList()) {
                if (!CanContainTemplateBlockMarkers(child)) {
                    continue;
                }

                generated += ExecuteRepeatingBlocks(child, regions, removeFields);
            }

            return generated;
        }

        private static int ExecuteConditionalBlocks(OpenXmlCompositeElement container, IDictionary<string, bool> conditions, bool removeMarkers) {
            var elements = container.ChildElements.ToList();
            var stack = new List<ConditionalBlockStart>();
            var blocks = new List<ConditionalBlockRange>();

            for (int index = 0; index < elements.Count; index++) {
                if (elements[index] is not Paragraph paragraph || !TryGetConditionalBlockMarker(paragraph, out string? markerName, out bool isStart)) {
                    continue;
                }

                if (isStart) {
                    stack.Add(new ConditionalBlockStart(markerName!, index));
                    continue;
                }

                if (stack.Count == 0) {
                    throw new InvalidOperationException($"Conditional block end marker '{{{{/{markerName}}}}}' does not have a matching start marker.");
                }

                var start = stack[stack.Count - 1];
                stack.RemoveAt(stack.Count - 1);
                if (!string.Equals(start.Name, markerName, StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidOperationException($"Conditional block end marker '{{{{/{markerName}}}}}' does not match start marker '{{{{#{start.Name}}}}}'.");
                }

                blocks.Add(new ConditionalBlockRange(start.Name, start.Index, index));
            }

            if (stack.Count > 0) {
                var start = stack[stack.Count - 1];
                throw new InvalidOperationException($"Conditional block start marker '{{{{#{start.Name}}}}}' does not have a matching end marker.");
            }

            foreach (var block in blocks) {
                if (!conditions.ContainsKey(block.Name)) {
                    throw new InvalidOperationException($"Conditional block '{block.Name}' was not supplied.");
                }
            }

            foreach (var block in blocks.OrderByDescending(item => item.StartIndex)) {
                bool include = conditions[block.Name];

                if (include) {
                    if (removeMarkers) {
                        RemoveIfAttached(elements[block.EndIndex]);
                        RemoveIfAttached(elements[block.StartIndex]);
                    }

                    continue;
                }

                for (int index = block.StartIndex; index <= block.EndIndex; index++) {
                    RemoveIfAttached(elements[index]);
                }
            }

            int processed = blocks.Count;
            foreach (var child in container.ChildElements.OfType<OpenXmlCompositeElement>().ToList()) {
                if (!CanContainTemplateBlockMarkers(child)) {
                    continue;
                }

                processed += ExecuteConditionalBlocks(child, conditions, removeMarkers);
            }

            return processed;
        }

        private static bool TryGetConditionalBlockMarker(Paragraph paragraph, out string? name, out bool isStart) {
            string text = paragraph.InnerText ?? string.Empty;
            var match = ConditionalBlockMarkerRegex.Match(text);
            if (!match.Success) {
                name = null;
                isStart = false;
                return false;
            }

            name = match.Groups["name"].Value;
            isStart = match.Groups["kind"].Value == "#";
            return true;
        }

        private static bool TryGetRepeatingBlockMarker(Paragraph paragraph, out string? name, out bool isStart) {
            string text = paragraph.InnerText ?? string.Empty;
            var match = RepeatingBlockMarkerRegex.Match(text);
            if (!match.Success) {
                name = null;
                isStart = false;
                return false;
            }

            name = match.Groups["name"].Value;
            isStart = match.Groups["kind"].Value.StartsWith("#", StringComparison.Ordinal);
            return true;
        }

        private static void RemoveIfAttached(OpenXmlElement element) {
            if (element.Parent != null) {
                element.Remove();
            }
        }

        private static void ReplaceSimpleMergeFields(OpenXmlElement root, IDictionary<string, string> values, bool removeFields) {
            foreach (var simpleField in root.Descendants<SimpleField>().ToList()) {
                string? name = GetMergeFieldName(simpleField.Instruction?.Value);
                if (name == null || !values.TryGetValue(name, out string? value)) {
                    continue;
                }

                if (removeFields) {
                    var replacement = CreateReplacementRun(value, simpleField.Elements<Run>().FirstOrDefault());
                    simpleField.InsertBeforeSelf(replacement);
                    simpleField.Remove();
                } else {
                    SetFieldResultText(simpleField.Elements<Run>(), value);
                }
            }
        }

        private static void ReplaceComplexMergeFields(OpenXmlElement root, IDictionary<string, string> values, bool removeFields) {
            foreach (var paragraph in EnumerateParagraphs(root)) {
                List<Run>? fieldRuns = null;

                foreach (var run in paragraph.Elements<Run>().ToList()) {
                    var fieldChar = run.Elements<FieldChar>().FirstOrDefault();
                    if (fieldChar?.FieldCharType?.Value == FieldCharValues.Begin) {
                        fieldRuns = new List<Run> { run };
                        continue;
                    }

                    if (fieldRuns == null) {
                        continue;
                    }

                    fieldRuns.Add(run);
                    if (fieldChar?.FieldCharType?.Value != FieldCharValues.End) {
                        continue;
                    }

                    ReplaceComplexFieldRuns(fieldRuns, values, removeFields);
                    fieldRuns = null;
                }
            }
        }

        private static IEnumerable<Paragraph> EnumerateParagraphs(OpenXmlElement root) {
            if (root is Paragraph paragraph) {
                yield return paragraph;
            }

            foreach (var child in root.Descendants<Paragraph>()) {
                yield return child;
            }
        }

        private static IEnumerable<OpenXmlCompositeElement> EnumerateTemplateRoots(WordDocument document) {
            MainDocumentPart? mainPart = document._wordprocessingDocument.MainDocumentPart;
            Body? body = mainPart?.Document?.Body;
            if (body != null) {
                yield return body;
            }

            if (mainPart == null) {
                yield break;
            }

            foreach (HeaderPart headerPart in mainPart.HeaderParts) {
                if (headerPart.Header != null) {
                    yield return headerPart.Header;
                }
            }

            foreach (FooterPart footerPart in mainPart.FooterParts) {
                if (footerPart.Footer != null) {
                    yield return footerPart.Footer;
                }
            }
        }

        private static void ReplaceComplexFieldRuns(IReadOnlyList<Run> fieldRuns, IDictionary<string, string> values, bool removeFields) {
            string instruction = string.Concat(fieldRuns
                .SelectMany(run => run.Elements<FieldCode>())
                .Select(code => code.Text));
            string? name = GetMergeFieldName(instruction);
            if (name == null || !values.TryGetValue(name, out string? value)) {
                return;
            }

            if (removeFields) {
                Run? sourceRun = GetComplexFieldResultRuns(fieldRuns).FirstOrDefault()
                    ?? fieldRuns.FirstOrDefault(run => run.GetFirstChild<RunProperties>() != null)
                    ?? fieldRuns.FirstOrDefault();
                var replacement = CreateReplacementRun(value, sourceRun);
                fieldRuns[0].InsertBeforeSelf(replacement);
                foreach (var run in fieldRuns) {
                    run.Remove();
                }

                return;
            }

            var resultRuns = GetComplexFieldResultRuns(fieldRuns).ToList();
            SetFieldResultText(resultRuns, value);
        }

        private static IEnumerable<Run> GetComplexFieldResultRuns(IReadOnlyList<Run> fieldRuns) {
            bool afterSeparator = false;

            foreach (var run in fieldRuns) {
                var fieldChar = run.Elements<FieldChar>().FirstOrDefault();
                if (fieldChar?.FieldCharType?.Value == FieldCharValues.Separate) {
                    afterSeparator = true;
                    continue;
                }

                if (fieldChar?.FieldCharType?.Value == FieldCharValues.End) {
                    yield break;
                }

                if (afterSeparator) {
                    yield return run;
                }
            }
        }

        private static void SetFieldResultText(IEnumerable<Run> runs, string value) {
            var textElements = runs
                .SelectMany(run => run.Elements<Text>())
                .ToList();

            if (textElements.Count == 0) {
                return;
            }

            textElements[0].Text = value;
            textElements[0].Space = SpaceProcessingModeValues.Preserve;
            for (int i = 1; i < textElements.Count; i++) {
                textElements[i].Text = string.Empty;
            }
        }

        private static Run CreateReplacementRun(string value, Run? sourceRun) {
            var run = new Run();
            var properties = sourceRun?.GetFirstChild<RunProperties>();
            if (properties != null) {
                run.Append((RunProperties)properties.CloneNode(true));
            }

            run.Append(new Text(value) { Space = SpaceProcessingModeValues.Preserve });
            return run;
        }

        private static string? GetMergeFieldName(string? fieldInstruction) {
            if (string.IsNullOrWhiteSpace(fieldInstruction)) {
                return null;
            }

            var parser = new WordFieldParser(fieldInstruction!);
            if (parser.WordFieldType != WordFieldType.MergeField || parser.Instructions.Count == 0) {
                return null;
            }

            return parser.Instructions[0].Trim().Trim('"');
        }

        private static IEnumerable<string> EnumerateMergeFieldNames(OpenXmlElement root) {
            foreach (var simpleField in root.Descendants<SimpleField>()) {
                string? name = TryGetMergeFieldName(simpleField.Instruction?.Value);
                if (!string.IsNullOrWhiteSpace(name)) {
                    yield return name!;
                }
            }

            foreach (var paragraph in EnumerateParagraphs(root)) {
                List<Run>? fieldRuns = null;
                foreach (var run in paragraph.Elements<Run>()) {
                    var fieldChar = run.Elements<FieldChar>().FirstOrDefault();
                    if (fieldChar?.FieldCharType?.Value == FieldCharValues.Begin) {
                        fieldRuns = new List<Run> { run };
                        continue;
                    }

                    if (fieldRuns == null) {
                        continue;
                    }

                    fieldRuns.Add(run);
                    if (fieldChar?.FieldCharType?.Value != FieldCharValues.End) {
                        continue;
                    }

                    string instruction = string.Concat(fieldRuns
                        .SelectMany(item => item.Elements<FieldCode>())
                        .Select(code => code.Text));
                    string? name = TryGetMergeFieldName(instruction);
                    if (!string.IsNullOrWhiteSpace(name)) {
                        yield return name!;
                    }

                    fieldRuns = null;
                }
            }
        }

        private static string? TryGetMergeFieldName(string? fieldInstruction) {
            try {
                return GetMergeFieldName(fieldInstruction);
            } catch (NotImplementedException) {
                return null;
            }
        }

        private static Dictionary<string, IEnumerable<WordMailMergeBlockData>> ConvertFlatRegions(IDictionary<string, IEnumerable<IDictionary<string, string>>> regions) {
            var converted = new Dictionary<string, IEnumerable<WordMailMergeBlockData>>(StringComparer.OrdinalIgnoreCase);
            foreach (var pair in regions) {
                converted[pair.Key] = pair.Value.Select(values => new WordMailMergeBlockData(values));
            }

            return converted;
        }

        private static bool TryGetRegionRows(IDictionary<string, IEnumerable<WordMailMergeBlockData>> regions, string name, out IEnumerable<WordMailMergeBlockData>? rows) {
            if (regions.TryGetValue(name, out rows)) {
                return true;
            }

            foreach (var pair in regions) {
                if (string.Equals(pair.Key, name, StringComparison.OrdinalIgnoreCase)) {
                    rows = pair.Value;
                    return true;
                }
            }

            rows = null;
            return false;
        }

        private static ConditionalBlockInspection InspectConditionalBlocks(OpenXmlCompositeElement? container) {
            if (container == null) {
                return new ConditionalBlockInspection(Array.Empty<string>(), Array.Empty<WordMailMergeTemplateIssue>());
            }

            var names = new List<string>();
            var issues = new List<WordMailMergeTemplateIssue>();
            InspectConditionalBlocks(container, names, issues);
            return new ConditionalBlockInspection(names, issues);
        }

        private static RepeatingBlockInspection InspectRepeatingBlocks(OpenXmlCompositeElement? container) {
            if (container == null) {
                return new RepeatingBlockInspection(Array.Empty<string>(), Array.Empty<WordMailMergeTemplateIssue>());
            }

            var names = new List<string>();
            var issues = new List<WordMailMergeTemplateIssue>();
            InspectRepeatingBlocks(container, names, issues);
            return new RepeatingBlockInspection(names, issues);
        }

        private static void InspectRepeatingBlocks(OpenXmlCompositeElement container, List<string> names, List<WordMailMergeTemplateIssue> issues) {
            var elements = container.ChildElements.ToList();
            var stack = new List<RepeatingBlockStart>();

            for (int index = 0; index < elements.Count; index++) {
                if (elements[index] is not Paragraph paragraph || !TryGetRepeatingBlockMarker(paragraph, out string? markerName, out bool isStart)) {
                    continue;
                }

                names.Add(markerName!);
                if (isStart) {
                    stack.Add(new RepeatingBlockStart(markerName!, index));
                    continue;
                }

                if (stack.Count == 0) {
                    issues.Add(new WordMailMergeTemplateIssue(
                        WordMailMergeTemplateIssueKind.UnmatchedRepeatingBlockEnd,
                        markerName!,
                        $"Repeating block end marker '{{{{/each {markerName}}}}}' does not have a matching start marker."));
                    continue;
                }

                var start = stack[stack.Count - 1];
                stack.RemoveAt(stack.Count - 1);
                if (!string.Equals(start.Name, markerName, StringComparison.OrdinalIgnoreCase)) {
                    issues.Add(new WordMailMergeTemplateIssue(
                        WordMailMergeTemplateIssueKind.MismatchedRepeatingBlockEnd,
                        markerName!,
                        $"Repeating block end marker '{{{{/each {markerName}}}}}' does not match start marker '{{{{#each {start.Name}}}}}'."));
                }
            }

            foreach (var start in stack) {
                issues.Add(new WordMailMergeTemplateIssue(
                    WordMailMergeTemplateIssueKind.UnmatchedRepeatingBlockStart,
                    start.Name,
                    $"Repeating block start marker '{{{{#each {start.Name}}}}}' does not have a matching end marker."));
            }

            foreach (var child in container.ChildElements.OfType<OpenXmlCompositeElement>().ToList()) {
                if (!CanContainTemplateBlockMarkers(child)) {
                    continue;
                }

                InspectRepeatingBlocks(child, names, issues);
            }
        }

        private static void InspectConditionalBlocks(OpenXmlCompositeElement container, List<string> names, List<WordMailMergeTemplateIssue> issues) {
            var elements = container.ChildElements.ToList();
            var stack = new List<ConditionalBlockStart>();

            for (int index = 0; index < elements.Count; index++) {
                if (elements[index] is not Paragraph paragraph || !TryGetConditionalBlockMarker(paragraph, out string? markerName, out bool isStart)) {
                    continue;
                }

                names.Add(markerName!);
                if (isStart) {
                    stack.Add(new ConditionalBlockStart(markerName!, index));
                    continue;
                }

                if (stack.Count == 0) {
                    issues.Add(new WordMailMergeTemplateIssue(
                        WordMailMergeTemplateIssueKind.UnmatchedConditionalEnd,
                        markerName!,
                        $"Conditional block end marker '{{{{/{markerName}}}}}' does not have a matching start marker."));
                    continue;
                }

                var start = stack[stack.Count - 1];
                stack.RemoveAt(stack.Count - 1);
                if (!string.Equals(start.Name, markerName, StringComparison.OrdinalIgnoreCase)) {
                    issues.Add(new WordMailMergeTemplateIssue(
                        WordMailMergeTemplateIssueKind.MismatchedConditionalEnd,
                        markerName!,
                        $"Conditional block end marker '{{{{/{markerName}}}}}' does not match start marker '{{{{#{start.Name}}}}}'."));
                }
            }

            foreach (var start in stack) {
                issues.Add(new WordMailMergeTemplateIssue(
                    WordMailMergeTemplateIssueKind.UnmatchedConditionalStart,
                    start.Name,
                    $"Conditional block start marker '{{{{#{start.Name}}}}}' does not have a matching end marker."));
            }

            foreach (var child in container.ChildElements.OfType<OpenXmlCompositeElement>().ToList()) {
                if (!CanContainTemplateBlockMarkers(child)) {
                    continue;
                }

                InspectConditionalBlocks(child, names, issues);
            }
        }

        private static bool CanContainTemplateBlockMarkers(OpenXmlCompositeElement element) {
            return element is not Paragraph && element is not Run;
        }

        private static IEnumerable<BoundContentControl> EnumerateBoundContentControls(MainDocumentPart mainPart) {
            foreach (OpenXmlPart part in EnumerateParts(mainPart)) {
                OpenXmlPartRootElement? root = GetPartRootElement(part);
                if (root == null) {
                    continue;
                }

                foreach (SdtRun control in root.Descendants<SdtRun>()) {
                    DataBinding? binding = control.SdtProperties?.GetFirstChild<DataBinding>();
                    if (binding != null) {
                        yield return new BoundContentControl(part, control, control.SdtProperties!, binding);
                    }
                }

                foreach (SdtBlock control in root.Descendants<SdtBlock>()) {
                    DataBinding? binding = control.SdtProperties?.GetFirstChild<DataBinding>();
                    if (binding != null) {
                        yield return new BoundContentControl(part, control, control.SdtProperties!, binding);
                    }
                }

                foreach (SdtCell control in root.Descendants<SdtCell>()) {
                    DataBinding? binding = control.SdtProperties?.GetFirstChild<DataBinding>();
                    if (binding != null) {
                        yield return new BoundContentControl(part, control, control.SdtProperties!, binding);
                    }
                }
            }
        }

        private static OpenXmlPartRootElement? GetPartRootElement(OpenXmlPart part) {
            switch (part) {
                case MainDocumentPart mainPart:
                    return mainPart.Document;
                case HeaderPart headerPart:
                    return headerPart.Header;
                case FooterPart footerPart:
                    return footerPart.Footer;
                case FootnotesPart footnotesPart:
                    return footnotesPart.Footnotes;
                case EndnotesPart endnotesPart:
                    return endnotesPart.Endnotes;
                case WordprocessingCommentsPart commentsPart:
                    return commentsPart.Comments;
                default:
                    return part.RootElement;
            }
        }

        private static IEnumerable<OpenXmlPart> EnumerateParts(OpenXmlPart root) {
            yield return root;

            foreach (IdPartPair pair in root.Parts) {
                OpenXmlPart part = pair.OpenXmlPart;
                foreach (OpenXmlPart child in EnumerateParts(part)) {
                    yield return child;
                }
            }
        }

        private static bool TryGetSuppliedBindingValue(IDictionary<string, string> values, BoundContentControl boundControl, out string? value) {
            foreach (string key in GetBindingKeys(boundControl)) {
                if (TryGetValueOrdinalIgnoreCase(values, key, out value)) {
                    return true;
                }
            }

            value = null;
            return false;
        }

        private static bool TryGetValueOrdinalIgnoreCase(IDictionary<string, string> values, string key, out string? value) {
            if (values.TryGetValue(key, out string? exactValue)) {
                value = exactValue;
                return true;
            }

            foreach (var pair in values) {
                if (string.Equals(pair.Key, key, StringComparison.OrdinalIgnoreCase)) {
                    value = pair.Value;
                    return true;
                }
            }

            value = null;
            return false;
        }

        private static IEnumerable<string> GetBindingKeys(BoundContentControl boundControl) {
            string? alias = boundControl.Properties.Descendants<SdtAlias>().FirstOrDefault()?.Val?.ToString();
            if (!string.IsNullOrWhiteSpace(alias)) {
                yield return alias!;
            }

            string? tag = boundControl.Properties.Descendants<Tag>().FirstOrDefault()?.Val?.ToString();
            if (!string.IsNullOrWhiteSpace(tag)) {
                yield return tag!;
            }

            string? xpath = boundControl.Binding.XPath?.Value;
            if (!string.IsNullOrWhiteSpace(xpath)) {
                yield return xpath!;

                string? storeItemId = boundControl.Binding.StoreItemId?.Value;
                if (!string.IsNullOrWhiteSpace(storeItemId)) {
                    yield return storeItemId + "|" + xpath;
                }
            }
        }

        private static string GetBindingDisplayKey(BoundContentControl boundControl) {
            return GetBindingKeys(boundControl).FirstOrDefault()
                ?? boundControl.Binding.XPath?.Value
                ?? boundControl.Binding.StoreItemId?.Value
                ?? boundControl.Part.Uri.OriginalString;
        }

        private static bool TryGetCustomXmlBindingValue(MainDocumentPart mainPart, DataBinding binding, Dictionary<CustomXmlPart, XDocument> cache, out string? value) {
            if (TrySelectCustomXmlBindingElement(mainPart, binding, cache, out _, out XElement? element)) {
                value = element.Value;
                return true;
            }

            value = null;
            return false;
        }

        private static bool TrySelectCustomXmlBindingElement(MainDocumentPart mainPart, DataBinding binding, Dictionary<CustomXmlPart, XDocument> cache, out CustomXmlPart customXmlPart, out XElement element) {
            string? storeItemId = binding.StoreItemId?.Value;
            IEnumerable<CustomXmlPart> parts = mainPart.CustomXmlParts;
            if (!string.IsNullOrWhiteSpace(storeItemId)) {
                var matchingParts = parts
                    .Where(part => string.Equals(GetCustomXmlStoreItemId(part), storeItemId, StringComparison.OrdinalIgnoreCase))
                    .ToList();
                if (matchingParts.Count > 0) {
                    parts = matchingParts;
                }
            }

            foreach (CustomXmlPart part in parts) {
                XDocument document = GetCustomXmlDocument(part, cache);
                XElement? selected = SelectBoundElement(document, binding);
                if (selected != null) {
                    customXmlPart = part;
                    element = selected;
                    return true;
                }
            }

            customXmlPart = null!;
            element = null!;
            return false;
        }

        private static string? GetCustomXmlStoreItemId(CustomXmlPart part) {
            return part.CustomXmlPropertiesPart?.DataStoreItem?.ItemId?.Value;
        }

        private static XDocument GetCustomXmlDocument(CustomXmlPart part, Dictionary<CustomXmlPart, XDocument> cache) {
            if (cache.TryGetValue(part, out XDocument? document)) {
                return document;
            }

            using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read)) {
                document = XDocument.Load(stream, LoadOptions.PreserveWhitespace);
            }

            cache.Add(part, document);
            return document;
        }

        private static XElement? SelectBoundElement(XDocument document, DataBinding binding) {
            string? xpath = binding.XPath?.Value;
            if (string.IsNullOrWhiteSpace(xpath)) {
                return null;
            }

            XmlNamespaceManager namespaceManager = CreateNamespaceManager(binding.PrefixMappings?.Value);
            return document.XPathSelectElement(xpath!, namespaceManager);
        }

        private static XmlNamespaceManager CreateNamespaceManager(string? prefixMappings) {
            var manager = new XmlNamespaceManager(new NameTable());
            if (string.IsNullOrWhiteSpace(prefixMappings)) {
                return manager;
            }

            foreach (Match match in Regex.Matches(prefixMappings!, @"xmlns(?::(?<prefix>[\w.-]+))?\s*=\s*(?<quote>['""])(?<uri>.*?)\k<quote>")) {
                string prefix = match.Groups["prefix"].Success ? match.Groups["prefix"].Value : string.Empty;
                string uri = match.Groups["uri"].Value;
                if (!string.IsNullOrWhiteSpace(uri)) {
                    manager.AddNamespace(prefix, uri);
                }
            }

            return manager;
        }

        private static void SaveCustomXmlPart(CustomXmlPart part, XDocument document) {
            using (Stream stream = part.GetStream(FileMode.Create, FileAccess.Write)) {
                document.Save(stream);
            }
        }

        private static bool UpdateContentControlText(OpenXmlElement contentControl, string value) {
            switch (contentControl) {
                case SdtRun run:
                    run.SdtContentRun ??= new SdtContentRun();
                    SetTextInComposite(run.SdtContentRun, value, () => {
                        var newRun = new Run(new Text { Space = SpaceProcessingModeValues.Preserve });
                        run.SdtContentRun.Append(newRun);
                        return newRun.GetFirstChild<Text>()!;
                    });
                    return true;
                case SdtBlock block:
                    block.SdtContentBlock ??= new SdtContentBlock();
                    SetTextInComposite(block.SdtContentBlock, value, () => {
                        var paragraph = new Paragraph(new Run(new Text { Space = SpaceProcessingModeValues.Preserve }));
                        block.SdtContentBlock.Append(paragraph);
                        return paragraph.Descendants<Text>().First();
                    });
                    return true;
                case SdtCell cell:
                    cell.SdtContentCell ??= new SdtContentCell();
                    SetTextInComposite(cell.SdtContentCell, value, () => {
                        var tableCell = new TableCell(new Paragraph(new Run(new Text { Space = SpaceProcessingModeValues.Preserve })));
                        cell.SdtContentCell.Append(tableCell);
                        return tableCell.Descendants<Text>().First();
                    });
                    return true;
                default:
                    return false;
            }
        }

        private static void SetTextInComposite(OpenXmlCompositeElement container, string value, Func<Text> createText) {
            var textElements = container.Descendants<Text>().ToList();
            Text firstText = textElements.FirstOrDefault() ?? createText();
            firstText.Text = value;
            firstText.Space = SpaceProcessingModeValues.Preserve;

            foreach (Text extraText in textElements.Skip(1)) {
                extraText.Text = string.Empty;
            }
        }

        private readonly struct ConditionalBlockStart {
            internal ConditionalBlockStart(string name, int index) {
                Name = name;
                Index = index;
            }

            internal string Name { get; }
            internal int Index { get; }
        }

        private readonly struct ConditionalBlockRange {
            internal ConditionalBlockRange(string name, int startIndex, int endIndex) {
                Name = name;
                StartIndex = startIndex;
                EndIndex = endIndex;
            }

            internal string Name { get; }
            internal int StartIndex { get; }
            internal int EndIndex { get; }
        }

        private readonly struct RepeatingBlockStart {
            internal RepeatingBlockStart(string name, int index) {
                Name = name;
                Index = index;
            }

            internal string Name { get; }
            internal int Index { get; }
        }

        private readonly struct RepeatingBlockRange {
            internal RepeatingBlockRange(string name, int startIndex, int endIndex) {
                Name = name;
                StartIndex = startIndex;
                EndIndex = endIndex;
            }

            internal string Name { get; }
            internal int StartIndex { get; }
            internal int EndIndex { get; }
        }

        private readonly struct ConditionalBlockInspection {
            internal ConditionalBlockInspection(IReadOnlyList<string> names, IReadOnlyList<WordMailMergeTemplateIssue> issues) {
                Names = names;
                Issues = issues;
            }

            internal IReadOnlyList<string> Names { get; }
            internal IReadOnlyList<WordMailMergeTemplateIssue> Issues { get; }
        }

        private readonly struct RepeatingBlockInspection {
            internal RepeatingBlockInspection(IReadOnlyList<string> names, IReadOnlyList<WordMailMergeTemplateIssue> issues) {
                Names = names;
                Issues = issues;
            }

            internal IReadOnlyList<string> Names { get; }
            internal IReadOnlyList<WordMailMergeTemplateIssue> Issues { get; }
        }

        private sealed class BoundContentControl {
            internal BoundContentControl(OpenXmlPart part, OpenXmlElement contentControl, SdtProperties properties, DataBinding binding) {
                Part = part;
                ContentControl = contentControl;
                Properties = properties;
                Binding = binding;
            }

            internal OpenXmlPart Part { get; }
            internal OpenXmlElement ContentControl { get; }
            internal SdtProperties Properties { get; }
            internal DataBinding Binding { get; }
        }
    }
}
