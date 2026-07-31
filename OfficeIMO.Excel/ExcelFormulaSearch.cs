using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Excel {
    /// <summary>Controls reference-aware formula search.</summary>
    public sealed class ExcelFormulaSearchOptions {
        /// <summary>Optional formula text fragment, excluding or including a leading equals sign.</summary>
        public string? Text { get; set; }

        /// <summary>Optional A1 cell/range reference. Intersecting formula references are matched.</summary>
        public string? Reference { get; set; }

        /// <summary>Optional function name such as <c>SUM</c> or <c>XLOOKUP</c>.</summary>
        public string? Function { get; set; }

        /// <summary>Whether text and function matching is case-sensitive.</summary>
        public bool MatchCase { get; set; }

        /// <summary>Maximum returned formula cells.</summary>
        public int MaximumResults { get; set; } = 1_000;

        internal void Validate() {
            if (string.IsNullOrWhiteSpace(Text)
                && string.IsNullOrWhiteSpace(Reference)
                && string.IsNullOrWhiteSpace(Function)) {
                throw new InvalidOperationException("Formula search requires Text, Reference, or Function.");
            }
            if (MaximumResults < 1) throw new ArgumentOutOfRangeException(nameof(MaximumResults));
            if (!string.IsNullOrWhiteSpace(Function)
                && Function!.Any(character => !(char.IsLetterOrDigit(character) || character == '_' || character == '.'))) {
                throw new ArgumentException("Formula function names may contain only letters, digits, underscores, and periods.", nameof(Function));
            }
        }
    }

    public partial class ExcelSheet {
        /// <summary>Searches formula text, function calls, or parsed references on this worksheet.</summary>
        public IReadOnlyList<ExcelFormulaCellInfo> SearchFormulas(ExcelFormulaSearchOptions options) =>
            SearchFormulaCells(GetFormulaCells(), options);

        internal static IReadOnlyList<ExcelFormulaCellInfo> SearchFormulaCells(
            IEnumerable<ExcelFormulaCellInfo> formulas,
            ExcelFormulaSearchOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            options.Validate();
            ExcelReference? target = null;
            if (!string.IsNullOrWhiteSpace(options.Reference)) {
                target = ExcelReference.Parse(options.Reference!);
            }

            var matches = new List<ExcelFormulaCellInfo>();
            StringComparison comparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
            foreach (ExcelFormulaCellInfo formula in formulas) {
                if (!string.IsNullOrWhiteSpace(options.Text)
                    && formula.Formula.IndexOf(options.Text!, comparison) < 0) {
                    continue;
                }
                if (!string.IsNullOrWhiteSpace(options.Function)
                    && !ContainsFunction(formula.Formula, options.Function!, comparison)) {
                    continue;
                }
                if (target != null && !ContainsIntersectingReference(formula, target)) {
                    continue;
                }

                matches.Add(formula);
                if (matches.Count >= options.MaximumResults) break;
            }
            return new ReadOnlyCollection<ExcelFormulaCellInfo>(matches);
        }

        private static bool ContainsIntersectingReference(ExcelFormulaCellInfo formula, ExcelReference target) {
            foreach (ExcelFormulaReferenceSyntax syntax in formula.SyntaxTree.Nodes.OfType<ExcelFormulaReferenceSyntax>()) {
                ExcelReference candidate = syntax.Reference;
                string candidateQualifier = NormalizeQualifier(candidate.Qualifier ?? formula.SheetName);
                string targetQualifier = NormalizeQualifier(target.Qualifier ?? formula.SheetName);
                if (!string.Equals(candidateQualifier, targetQualifier, StringComparison.OrdinalIgnoreCase)) continue;
                candidate.GetBounds(out int cr1, out int cc1, out int cr2, out int cc2);
                target.GetBounds(out int tr1, out int tc1, out int tr2, out int tc2);
                if (cr1 <= tr2 && cr2 >= tr1 && cc1 <= tc2 && cc2 >= tc1) return true;
            }
            return false;
        }

        private static bool ContainsFunction(string formula, string requested, StringComparison comparison) {
            int index = 0;
            bool inString = false;
            while (index < formula.Length) {
                char current = formula[index];
                if (current == '"') {
                    if (inString && index + 1 < formula.Length && formula[index + 1] == '"') { index += 2; continue; }
                    inString = !inString;
                    index++;
                    continue;
                }
                if (inString || !(char.IsLetter(current) || current == '_')) { index++; continue; }

                int start = index++;
                while (index < formula.Length
                    && (char.IsLetterOrDigit(formula[index]) || formula[index] == '_' || formula[index] == '.')) index++;
                int cursor = index;
                while (cursor < formula.Length && char.IsWhiteSpace(formula[cursor])) cursor++;
                if (cursor < formula.Length && formula[cursor] == '(') {
                    string name = NormalizeFunctionName(formula.Substring(start, index - start));
                    if (string.Equals(name, NormalizeFunctionName(requested), comparison)) return true;
                }
            }
            return false;
        }

        private static string NormalizeFunctionName(string value) {
            string result = value;
            if (result.StartsWith("_xlfn.", StringComparison.OrdinalIgnoreCase)) result = result.Substring(6);
            if (result.StartsWith("_xlws.", StringComparison.OrdinalIgnoreCase)) result = result.Substring(6);
            return result;
        }

        private static string NormalizeQualifier(string value) {
            string result = value.Trim();
            if (result.Length >= 2 && result[0] == '\'' && result[result.Length - 1] == '\'') {
                result = result.Substring(1, result.Length - 2).Replace("''", "'");
            }
            int workbookEnd = result.IndexOf(']');
            if (result.StartsWith("[", StringComparison.Ordinal) && workbookEnd >= 0) result = result.Substring(workbookEnd + 1);
            return result;
        }
    }

    public partial class ExcelDocument {
        /// <summary>Searches formula text, function calls, or parsed references across the workbook.</summary>
        public IReadOnlyList<ExcelFormulaCellInfo> SearchFormulas(ExcelFormulaSearchOptions options) =>
            ExcelSheet.SearchFormulaCells(InspectFormulas().Formulas, options);
    }
}
