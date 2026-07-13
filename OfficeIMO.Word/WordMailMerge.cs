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
    /// Provides basic mail merge capabilities by replacing <c>MERGEFIELD</c> fields with supplied values.
    /// </summary>
    public static partial class WordMailMerge {
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
                    document.Save(outputPath);
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

            var conditionLookup = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, bool> condition in conditions) {
                if (!string.IsNullOrWhiteSpace(condition.Key)) {
                    conditionLookup[condition.Key] = condition.Value;
                }
            }

            int processed = 0;
            foreach (var root in EnumerateTemplateRoots(document)) {
                processed += ExecuteConditionalBlocks(root, conditionLookup, removeMarkers);
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
                issues.AddRange(EnumerateUnsupportedMailMergeControlFieldIssues(root));
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
        /// Inspects a Word mail-merge template and returns a CI-friendly preflight report with capability checks and diagnostics.
        /// </summary>
        /// <param name="document">Document to inspect.</param>
        /// <param name="mergeFieldNames">Optional supplied MERGEFIELD names. When provided, missing fields are reported as issues.</param>
        /// <param name="conditionNames">Optional supplied conditional block names. When provided, missing conditions are reported as issues.</param>
        /// <param name="repeatingBlockNames">Optional supplied repeated block names. When provided, missing repeated block rows are reported as issues.</param>
        public static WordTemplatePreflightReport PreflightTemplate(WordDocument document, IEnumerable<string>? mergeFieldNames = null, IEnumerable<string>? conditionNames = null, IEnumerable<string>? repeatingBlockNames = null) {
            return WordTemplatePreflightReport.FromInspection(InspectTemplate(document, mergeFieldNames, conditionNames, repeatingBlockNames));
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

    }
}
