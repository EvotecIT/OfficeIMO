using System;
using System.Linq;
using SixLabors.ImageSharp;

namespace OfficeIMO.Word {
    public partial class WordTable {
        /// <summary>
        /// Applies conditional formatting to cells based on text contained in a specific column.
        /// </summary>
        /// <param name="columnName">Header text of the column to evaluate.</param>
        /// <param name="matchText">Text to compare against cell content.</param>
        /// <param name="matchType">Comparison operator.</param>
        /// <param name="matchFillColorHex">Background color applied when condition is met.</param>
        /// <param name="matchFontColorHex">Font color applied when condition is met.</param>
        /// <param name="noMatchFillColorHex">Background color applied when condition is not met.</param>
        /// <param name="noMatchFontColorHex">Font color applied when condition is not met.</param>
        /// <param name="ignoreCase">Whether comparison should ignore case.</param>
        /// <param name="highlightColumns">Additional columns to apply the formatting to.</param>
        /// <param name="matchTextFormat">Optional action applied to paragraphs when the condition is met.</param>
        /// <param name="noMatchTextFormat">Optional action applied to paragraphs when the condition is not met.</param>
        public void ConditionalFormatting(string columnName, string? matchText, TextMatchType matchType,
            string? matchFillColorHex = null, string? matchFontColorHex = null,
            string? noMatchFillColorHex = null, string? noMatchFontColorHex = null,
            bool ignoreCase = true, System.Collections.Generic.IEnumerable<string>? highlightColumns = null,
            Action<WordParagraph>? matchTextFormat = null, Action<WordParagraph>? noMatchTextFormat = null) {
            ArgumentException.ThrowIfNullOrEmpty(columnName);
            matchText ??= string.Empty;

            matchFillColorHex = matchFillColorHex != null ? Helpers.NormalizeColor(matchFillColorHex) : null;
            matchFontColorHex = matchFontColorHex != null ? Helpers.NormalizeColor(matchFontColorHex) : null;
            noMatchFillColorHex = noMatchFillColorHex != null ? Helpers.NormalizeColor(noMatchFillColorHex) : null;
            noMatchFontColorHex = noMatchFontColorHex != null ? Helpers.NormalizeColor(noMatchFontColorHex) : null;

            int headerIndex = -1;
            System.Collections.Generic.Dictionary<string, int> headerMap = new System.Collections.Generic.Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);
            if (Rows.Count > 0) {
                var headerRow = Rows[0];
                for (int i = 0; i < headerRow.CellsCount; i++) {
                    var text = headerRow.Cells[i].Paragraphs.FirstOrDefault()?.Text ?? string.Empty;
                    headerMap[text] = i;
                    if (string.Equals(text, columnName, ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal)) {
                        headerIndex = i;
                    }
                }
            }
            if (headerIndex == -1) {
                throw new ArgumentException($"Column '{columnName}' was not found.", nameof(columnName));
            }

            var highlightIndices = new System.Collections.Generic.HashSet<int> { headerIndex };
            if (highlightColumns != null) {
                foreach (var col in highlightColumns) {
                    if (string.IsNullOrEmpty(col)) continue;
                    if (headerMap.TryGetValue(col, out int idx)) {
                        highlightIndices.Add(idx);
                    }
                }
            }

            for (int r = 1; r < Rows.Count; r++) {
                var cell = Rows[r].Cells[headerIndex];
                var text = cell.Paragraphs.FirstOrDefault()?.Text ?? string.Empty;
                bool isMatch = matchType switch {
                    TextMatchType.Equals => string.Equals(text, matchText, ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal),
                    TextMatchType.Contains => text.IndexOf(matchText, ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal) >= 0,
                    TextMatchType.StartsWith => text.StartsWith(matchText, ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal),
                    TextMatchType.EndsWith => text.EndsWith(matchText, ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal),
                    _ => false
                };

                foreach (int colIndex in highlightIndices) {
                    var targetCell = Rows[r].Cells[colIndex];
                    if (isMatch) {
                        if (!string.IsNullOrEmpty(matchFillColorHex)) {
                            targetCell.ShadingFillColorHex = matchFillColorHex;
                        }
                    } else {
                        if (!string.IsNullOrEmpty(noMatchFillColorHex)) {
                            targetCell.ShadingFillColorHex = noMatchFillColorHex;
                        }
                    }

                    foreach (var p in targetCell.Paragraphs) {
                        if (isMatch) {
                            if (!string.IsNullOrEmpty(matchFontColorHex)) {
                                p.ColorHex = matchFontColorHex;
                            }
                            matchTextFormat?.Invoke(p);
                        } else {
                            if (!string.IsNullOrEmpty(noMatchFontColorHex)) {
                                p.ColorHex = noMatchFontColorHex;
                            }
                            noMatchTextFormat?.Invoke(p);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Applies conditional formatting to cells using <see cref="Color"/> parameters.
        /// </summary>
        /// <param name="columnName">Header text of the column to evaluate.</param>
        /// <param name="matchText">Text to compare against cell content.</param>
        /// <param name="matchType">Comparison operator.</param>
        /// <param name="matchFillColor">Background color applied when condition is met.</param>
        /// <param name="matchFontColor">Font color applied when condition is met.</param>
        /// <param name="noMatchFillColor">Background color applied when condition is not met.</param>
        /// <param name="noMatchFontColor">Font color applied when condition is not met.</param>
        /// <param name="ignoreCase">Whether comparison should ignore case.</param>
        /// <param name="highlightColumns">Additional columns to apply the formatting to.</param>
        /// <param name="matchTextFormat">Optional action applied to paragraphs when the condition is met.</param>
        /// <param name="noMatchTextFormat">Optional action applied to paragraphs when the condition is not met.</param>
        public void ConditionalFormatting(string columnName, string? matchText, TextMatchType matchType,
            Color matchFillColor, Color? matchFontColor = null,
            Color? noMatchFillColor = null, Color? noMatchFontColor = null,
            bool ignoreCase = true, System.Collections.Generic.IEnumerable<string>? highlightColumns = null,
            Action<WordParagraph>? matchTextFormat = null, Action<WordParagraph>? noMatchTextFormat = null) =>
            ConditionalFormatting(columnName, matchText, matchType,
                matchFillColorHex: matchFillColor.ToHexColor(),
                matchFontColorHex: matchFontColor?.ToHexColor(),
                noMatchFillColorHex: noMatchFillColor?.ToHexColor(),
                noMatchFontColorHex: noMatchFontColor?.ToHexColor(),
                ignoreCase: ignoreCase, highlightColumns: highlightColumns,
                matchTextFormat: matchTextFormat, noMatchTextFormat: noMatchTextFormat);

        /// <summary>
        /// Applies conditional formatting based on values in multiple columns.
        /// </summary>
        /// <param name="conditions">List of column/value conditions.</param>
        /// <param name="matchAll">When true, all conditions must match; otherwise any condition can match.</param>
        /// <param name="matchFillColorHex">Background color applied when conditions match.</param>
        /// <param name="matchFontColorHex">Font color applied when conditions match.</param>
        /// <param name="noMatchFillColorHex">Background color applied when conditions do not match.</param>
        /// <param name="noMatchFontColorHex">Font color applied when conditions do not match.</param>
        /// <param name="ignoreCase">Whether comparison should ignore case.</param>
        /// <param name="highlightColumns">Columns to apply the formatting to. Defaults to the columns used in conditions.</param>
        /// <param name="matchTextFormat">Optional action applied to paragraphs when conditions match.</param>
        /// <param name="noMatchTextFormat">Optional action applied to paragraphs when conditions do not match.</param>
        public void ConditionalFormatting(System.Collections.Generic.IEnumerable<(string ColumnName, string MatchText, TextMatchType MatchType)> conditions,
            bool matchAll,
            string? matchFillColorHex = null, string? matchFontColorHex = null,
            string? noMatchFillColorHex = null, string? noMatchFontColorHex = null,
            bool ignoreCase = true, System.Collections.Generic.IEnumerable<string>? highlightColumns = null,
            Action<WordParagraph>? matchTextFormat = null, Action<WordParagraph>? noMatchTextFormat = null) {
            ArgumentNullException.ThrowIfNull(conditions);

            var conditionList = conditions.ToList();
            if (conditionList.Count == 0) throw new ArgumentException("At least one condition is required", nameof(conditions));

            matchFillColorHex = matchFillColorHex != null ? Helpers.NormalizeColor(matchFillColorHex) : null;
            matchFontColorHex = matchFontColorHex != null ? Helpers.NormalizeColor(matchFontColorHex) : null;
            noMatchFillColorHex = noMatchFillColorHex != null ? Helpers.NormalizeColor(noMatchFillColorHex) : null;
            noMatchFontColorHex = noMatchFontColorHex != null ? Helpers.NormalizeColor(noMatchFontColorHex) : null;

            var headerMap = new System.Collections.Generic.Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);
            if (Rows.Count > 0) {
                var headerRow = Rows[0];
                for (int i = 0; i < headerRow.CellsCount; i++) {
                    var text = headerRow.Cells[i].Paragraphs.FirstOrDefault()?.Text ?? string.Empty;
                    headerMap[text] = i;
                }
            }
            var condIndices = new System.Collections.Generic.List<(int Index, string MatchText, TextMatchType Type)>();
            foreach (var c in conditionList) {
                if (!headerMap.TryGetValue(c.ColumnName, out int idx)) {
                    throw new ArgumentException($"Column '{c.ColumnName}' was not found.", nameof(conditions));
                }
                condIndices.Add((idx, c.MatchText ?? string.Empty, c.MatchType));
            }

            System.Collections.Generic.HashSet<int> highlightIndices;
            if (highlightColumns != null) {
                highlightIndices = new System.Collections.Generic.HashSet<int>();
                foreach (var col in highlightColumns) {
                    if (string.IsNullOrEmpty(col)) continue;
                    if (headerMap.TryGetValue(col, out int idx)) {
                        highlightIndices.Add(idx);
                    }
                }
                if (highlightIndices.Count == 0) {
                    highlightIndices.UnionWith(condIndices.Select(ci => ci.Index));
                }
            } else {
                highlightIndices = new System.Collections.Generic.HashSet<int>(condIndices.Select(ci => ci.Index));
            }

            static bool Check(string cellText, string matchText, TextMatchType type, bool ic) => type switch {
                TextMatchType.Equals => string.Equals(cellText, matchText, ic ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal),
                TextMatchType.Contains => cellText.IndexOf(matchText, ic ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal) >= 0,
                TextMatchType.StartsWith => cellText.StartsWith(matchText, ic ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal),
                TextMatchType.EndsWith => cellText.EndsWith(matchText, ic ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal),
                _ => false
            };

            for (int r = 1; r < Rows.Count; r++) {
                bool match = matchAll ? true : false;
                foreach (var cond in condIndices) {
                    var text = Rows[r].Cells[cond.Index].Paragraphs.FirstOrDefault()?.Text ?? string.Empty;
                    bool condRes = Check(text, cond.MatchText, cond.Type, ignoreCase);
                    if (matchAll) {
                        match &= condRes;
                        if (!match) break;
                    } else {
                        match |= condRes;
                        if (match) break;
                    }
                }

                foreach (int colIndex in highlightIndices) {
                    var targetCell = Rows[r].Cells[colIndex];
                    if (match) {
                        if (!string.IsNullOrEmpty(matchFillColorHex)) {
                            targetCell.ShadingFillColorHex = matchFillColorHex;
                        }
                    } else {
                        if (!string.IsNullOrEmpty(noMatchFillColorHex)) {
                            targetCell.ShadingFillColorHex = noMatchFillColorHex;
                        }
                    }

                    foreach (var p in targetCell.Paragraphs) {
                        if (match) {
                            if (!string.IsNullOrEmpty(matchFontColorHex)) {
                                p.ColorHex = matchFontColorHex;
                            }
                            matchTextFormat?.Invoke(p);
                        } else {
                            if (!string.IsNullOrEmpty(noMatchFontColorHex)) {
                                p.ColorHex = noMatchFontColorHex;
                            }
                            noMatchTextFormat?.Invoke(p);
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Applies conditional formatting based on values in multiple columns using <see cref="Color"/> parameters.
        /// </summary>
        /// <param name="conditions">List of column/value conditions.</param>
        /// <param name="matchAll">When true, all conditions must match; otherwise any condition can match.</param>
        /// <param name="matchFillColor">Background color applied when conditions match.</param>
        /// <param name="matchFontColor">Font color applied when conditions match.</param>
        /// <param name="noMatchFillColor">Background color applied when conditions do not match.</param>
        /// <param name="noMatchFontColor">Font color applied when conditions do not match.</param>
        /// <param name="ignoreCase">Whether comparison should ignore case.</param>
        /// <param name="highlightColumns">Columns to apply the formatting to. Defaults to the columns used in conditions.</param>
        /// <param name="matchTextFormat">Optional action applied to paragraphs when conditions match.</param>
        /// <param name="noMatchTextFormat">Optional action applied to paragraphs when conditions do not match.</param>
        public void ConditionalFormatting(System.Collections.Generic.IEnumerable<(string ColumnName, string MatchText, TextMatchType MatchType)> conditions,
            bool matchAll,
            Color matchFillColor, Color? matchFontColor = null,
            Color? noMatchFillColor = null, Color? noMatchFontColor = null,
            bool ignoreCase = true, System.Collections.Generic.IEnumerable<string>? highlightColumns = null,
            Action<WordParagraph>? matchTextFormat = null, Action<WordParagraph>? noMatchTextFormat = null) =>
            ConditionalFormatting(conditions, matchAll,
                matchFillColorHex: matchFillColor.ToHexColor(),
                matchFontColorHex: matchFontColor?.ToHexColor(),
                noMatchFillColorHex: noMatchFillColor?.ToHexColor(),
                noMatchFontColorHex: noMatchFontColor?.ToHexColor(),
                ignoreCase: ignoreCase, highlightColumns: highlightColumns,
                matchTextFormat: matchTextFormat, noMatchTextFormat: noMatchTextFormat);

        /// <summary>
        /// Applies conditional formatting based on a custom row predicate.
        /// </summary>
        /// <param name="predicate">Condition evaluated for each data row.</param>
        /// <param name="matchAction">Action executed when predicate is true.</param>
        /// <param name="noMatchAction">Action executed when predicate is false.</param>
        public void ConditionalFormatting(Func<WordTableRow, bool> predicate, Action<WordTableRow> matchAction, Action<WordTableRow>? noMatchAction = null) {
            ArgumentNullException.ThrowIfNull(predicate);
            ArgumentNullException.ThrowIfNull(matchAction);

            for (int r = 1; r < Rows.Count; r++) {
                var row = Rows[r];
                if (predicate(row)) {
                    matchAction(row);
                } else {
                    noMatchAction?.Invoke(row);
                }
            }
        }

        /// <summary>
        /// Starts building conditional formatting rules using a fluent API.
        /// </summary>
        /// <returns>A new <see cref="WordTableConditionalFormattingBuilder"/>.</returns>
        public WordTableConditionalFormattingBuilder BeginConditionalFormatting() => new WordTableConditionalFormattingBuilder(this);
    }

    /// <summary>
    /// Provides a fluent API for defining and applying multiple conditional formatting rules.
    /// </summary>
    public class WordTableConditionalFormattingBuilder {
        private readonly WordTable _table;
        private readonly System.Collections.Generic.List<System.Action> _rules = new();

        internal WordTableConditionalFormattingBuilder(WordTable table) {
            _table = table ?? throw new System.ArgumentNullException(nameof(table));
        }

        /// <summary>
        /// Adds a conditional formatting rule based on a column value.
        /// </summary>
        public WordTableConditionalFormattingBuilder AddRule(string columnName, string? matchText, TextMatchType matchType,
            string? matchFillColorHex = null, string? matchFontColorHex = null,
            string? noMatchFillColorHex = null, string? noMatchFontColorHex = null,
            bool ignoreCase = true, System.Collections.Generic.IEnumerable<string>? highlightColumns = null,
            Action<WordParagraph>? matchTextFormat = null, Action<WordParagraph>? noMatchTextFormat = null) {
            _rules.Add(() => _table.ConditionalFormatting(columnName, matchText, matchType,
                matchFillColorHex, matchFontColorHex, noMatchFillColorHex, noMatchFontColorHex,
                ignoreCase, highlightColumns,
                matchTextFormat, noMatchTextFormat));
            return this;
        }

        /// <summary>
        /// Adds a conditional formatting rule using <see cref="Color"/> parameters.
        /// </summary>
        public WordTableConditionalFormattingBuilder AddRule(string columnName, string? matchText, TextMatchType matchType,
            Color matchFillColor, Color? matchFontColor = null,
            Color? noMatchFillColor = null, Color? noMatchFontColor = null,
            bool ignoreCase = true, System.Collections.Generic.IEnumerable<string>? highlightColumns = null,
            Action<WordParagraph>? matchTextFormat = null, Action<WordParagraph>? noMatchTextFormat = null) {
            _rules.Add(() => _table.ConditionalFormatting(columnName, matchText, matchType,
                matchFillColor, matchFontColor, noMatchFillColor, noMatchFontColor,
                ignoreCase, highlightColumns,
                matchTextFormat, noMatchTextFormat));
            return this;
        }

        /// <summary>
        /// Adds a conditional formatting rule based on multiple column values.
        /// </summary>
        public WordTableConditionalFormattingBuilder AddRule(System.Collections.Generic.IEnumerable<(string ColumnName, string MatchText, TextMatchType MatchType)> conditions,
            bool matchAll,
            string? matchFillColorHex = null, string? matchFontColorHex = null,
            string? noMatchFillColorHex = null, string? noMatchFontColorHex = null,
            bool ignoreCase = true, System.Collections.Generic.IEnumerable<string>? highlightColumns = null,
            Action<WordParagraph>? matchTextFormat = null, Action<WordParagraph>? noMatchTextFormat = null) {
            _rules.Add(() => _table.ConditionalFormatting(conditions, matchAll,
                matchFillColorHex, matchFontColorHex, noMatchFillColorHex, noMatchFontColorHex,
                ignoreCase, highlightColumns,
                matchTextFormat, noMatchTextFormat));
            return this;
        }

        /// <summary>
        /// Adds a conditional formatting rule based on multiple column values using <see cref="Color"/> parameters.
        /// </summary>
        public WordTableConditionalFormattingBuilder AddRule(System.Collections.Generic.IEnumerable<(string ColumnName, string MatchText, TextMatchType MatchType)> conditions,
            bool matchAll,
            Color matchFillColor, Color? matchFontColor = null,
            Color? noMatchFillColor = null, Color? noMatchFontColor = null,
            bool ignoreCase = true, System.Collections.Generic.IEnumerable<string>? highlightColumns = null,
            Action<WordParagraph>? matchTextFormat = null, Action<WordParagraph>? noMatchTextFormat = null) {
            _rules.Add(() => _table.ConditionalFormatting(conditions, matchAll,
                matchFillColor, matchFontColor, noMatchFillColor, noMatchFontColor,
                ignoreCase, highlightColumns,
                matchTextFormat, noMatchTextFormat));
            return this;
        }

        /// <summary>
        /// Adds a conditional formatting rule using a predicate.
        /// </summary>
        public WordTableConditionalFormattingBuilder AddRule(System.Func<WordTableRow, bool> predicate, System.Action<WordTableRow> matchAction, System.Action<WordTableRow>? noMatchAction = null) {
            _rules.Add(() => _table.ConditionalFormatting(predicate, matchAction, noMatchAction));
            return this;
        }

        /// <summary>
        /// Applies all configured rules to the table.
        /// </summary>
        public void Apply() {
            foreach (var rule in _rules) {
                rule();
            }
        }
    }
}
