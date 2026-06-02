using System;
using System.Collections.Generic;
using System.Linq;

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
}
