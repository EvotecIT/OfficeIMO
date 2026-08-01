using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word {
    /// <summary>Outcome for one MERGEFIELD occurrence processed by OfficeIMO.</summary>
    public enum WordMailMergeFieldStatus {
        /// <summary>The supplied value was formatted and written.</summary>
        Merged,
        /// <summary>No supplied value matched the field name.</summary>
        MissingValue,
        /// <summary>The field requested formatting outside the deterministic OfficeIMO profile.</summary>
        UnsupportedFormatting,
        /// <summary>The instruction declared MERGEFIELD but could not be parsed as a named field.</summary>
        MalformedField
    }

    /// <summary>Structured result for one MERGEFIELD occurrence.</summary>
    public sealed class WordMailMergeFieldResult {
        internal WordMailMergeFieldResult(string name, string instruction, WordMailMergeFieldStatus status, string? value, string message) {
            Name = name;
            Instruction = instruction;
            Status = status;
            Value = value;
            Message = message;
        }

        /// <summary>Gets the merge-field name.</summary>
        public string Name { get; }
        /// <summary>Gets the original field instruction.</summary>
        public string Instruction { get; }
        /// <summary>Gets the merge outcome.</summary>
        public WordMailMergeFieldStatus Status { get; }
        /// <summary>Gets the formatted value when one was written.</summary>
        public string? Value { get; }
        /// <summary>Gets a stable human-readable diagnostic.</summary>
        public string Message { get; }
    }

    /// <summary>Summarizes a MERGEFIELD execution pass across body, headers, footers, notes, tables, content controls, and text boxes.</summary>
    public sealed class WordMailMergeExecutionReport {
        internal WordMailMergeExecutionReport(IReadOnlyList<WordMailMergeFieldResult> fields) {
            Fields = fields.ToArray();
        }

        /// <summary>Gets all field results in deterministic document-root order.</summary>
        public IReadOnlyList<WordMailMergeFieldResult> Fields { get; }
        /// <summary>Gets the number of merged fields.</summary>
        public int MergedCount => Fields.Count(item => item.Status == WordMailMergeFieldStatus.Merged);
        /// <summary>Gets unique missing value names.</summary>
        public IReadOnlyList<string> MissingValueNames => Fields
            .Where(item => item.Status == WordMailMergeFieldStatus.MissingValue)
            .Select(item => item.Name)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        /// <summary>Gets whether every discovered MERGEFIELD was merged.</summary>
        public bool IsComplete => Fields.All(item => item.Status == WordMailMergeFieldStatus.Merged);

        /// <summary>Throws with field diagnostics when the merge was incomplete; otherwise returns this report.</summary>
        public WordMailMergeExecutionReport EnsureComplete() {
            if (!IsComplete) {
                throw new InvalidOperationException(string.Join(Environment.NewLine,
                    Fields.Where(item => item.Status != WordMailMergeFieldStatus.Merged).Select(item => item.Message)));
            }
            return this;
        }
    }

    /// <summary>Outcome for one record emitted by a batch mail merge.</summary>
    public sealed class WordMailMergeBatchItemResult {
        internal WordMailMergeBatchItemResult(int recordIndex, string outputPath, WordMailMergeExecutionReport execution) {
            RecordIndex = recordIndex;
            OutputPath = outputPath;
            Execution = execution;
        }

        /// <summary>Gets the zero-based source record index.</summary>
        public int RecordIndex { get; }
        /// <summary>Gets the committed output path.</summary>
        public string OutputPath { get; }
        /// <summary>Gets per-field formatting and missing-value diagnostics.</summary>
        public WordMailMergeExecutionReport Execution { get; }
    }

    /// <summary>Summarizes a batch mail-merge operation in record order.</summary>
    public sealed class WordMailMergeBatchResult {
        internal WordMailMergeBatchResult(IReadOnlyList<WordMailMergeBatchItemResult> items) {
            Items = items.ToArray();
        }

        /// <summary>Gets all committed outputs and their field reports.</summary>
        public IReadOnlyList<WordMailMergeBatchItemResult> Items { get; }
        /// <summary>Gets committed output paths in record order.</summary>
        public IReadOnlyList<string> OutputPaths => Items.Select(item => item.OutputPath).ToArray();
        /// <summary>Gets whether every output merged every discovered field.</summary>
        public bool IsComplete => Items.All(item => item.Execution.IsComplete);

        /// <summary>Throws when any record has missing values, malformed fields, or unsupported formatting.</summary>
        public WordMailMergeBatchResult EnsureComplete() {
            foreach (WordMailMergeBatchItemResult item in Items) item.Execution.EnsureComplete();
            return this;
        }
    }

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
        MismatchedRepeatingBlockEnd,

        /// <summary>A Word-native mail-merge control field was found that OfficeIMO does not execute.</summary>
        UnsupportedMailMergeControlField,

        /// <summary>A MERGEFIELD requests formatting outside the deterministic OfficeIMO profile.</summary>
        UnsupportedMergeFieldFormatting,

        /// <summary>A MERGEFIELD contains a nested field and cannot be processed deterministically.</summary>
        MalformedMergeField
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
