using System.Text;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Workbook or worksheet formula inspection result.
    /// </summary>
    public sealed class ExcelFormulaInspection {
        internal ExcelFormulaInspection(IReadOnlyList<ExcelFormulaCellInfo> formulas) {
            Formulas = formulas;
            TotalFormulas = formulas.Count;
            SupportedFormulas = formulas.Count(formula => formula.IsSupportedByOfficeIMO);
            UnsupportedFormulas = TotalFormulas - SupportedFormulas;
            MissingCachedResults = formulas.Count(formula => !formula.HasCachedValue);
            DirtyFormulas = formulas.Count(formula => formula.IsDirty);
        }

        /// <summary>Formula cells discovered in workbook order.</summary>
        public IReadOnlyList<ExcelFormulaCellInfo> Formulas { get; }

        /// <summary>Total formula count.</summary>
        public int TotalFormulas { get; }

        /// <summary>Formula count supported by OfficeIMO's lightweight evaluator.</summary>
        public int SupportedFormulas { get; }

        /// <summary>Formula count that must be preserved, cached, or recalculated by Excel.</summary>
        public int UnsupportedFormulas { get; }

        /// <summary>Formula cells without cached results.</summary>
        public int MissingCachedResults { get; }

        /// <summary>Formula cells marked dirty for recalculation.</summary>
        public int DirtyFormulas { get; }

        /// <summary>True when every formula can be evaluated by OfficeIMO's lightweight evaluator.</summary>
        public bool AllSupported => TotalFormulas == SupportedFormulas;

        /// <summary>True when every formula has a cached result.</summary>
        public bool AllHaveCachedResults => MissingCachedResults == 0;

        /// <summary>Describes the formula patterns supported by OfficeIMO's lightweight evaluator.</summary>
        public ExcelFormulaCapabilities Capabilities => ExcelFormulaCapabilities.Current;

        /// <summary>
        /// Throws when any formula is outside OfficeIMO's lightweight evaluator support.
        /// </summary>
        public ExcelFormulaInspection EnsureAllSupported() {
            if (UnsupportedFormulas > 0) {
                var unsupported = Formulas
                    .Where(formula => !formula.IsSupportedByOfficeIMO)
                    .Select(formula => $"{formula.SheetName}!{formula.CellReference}")
                    .ToArray();
                throw new InvalidOperationException("Unsupported formulas: " + string.Join(", ", unsupported));
            }

            return this;
        }

        /// <summary>
        /// Throws when any formula lacks a cached result.
        /// </summary>
        public ExcelFormulaInspection EnsureAllHaveCachedResults() {
            if (MissingCachedResults > 0) {
                var missing = Formulas
                    .Where(formula => !formula.HasCachedValue)
                    .Select(formula => $"{formula.SheetName}!{formula.CellReference}")
                    .ToArray();
                throw new InvalidOperationException("Formula cells without cached results: " + string.Join(", ", missing));
            }

            return this;
        }

        /// <summary>
        /// Returns a compact Markdown report of formula support and cache status.
        /// </summary>
        public string ToMarkdown() {
            var builder = new StringBuilder();
            builder.AppendLine("# Excel Formula Inspection");
            builder.AppendLine();
            builder.AppendLine($"Total formulas: {TotalFormulas}");
            builder.AppendLine($"Supported formulas: {SupportedFormulas}");
            builder.AppendLine($"Unsupported formulas: {UnsupportedFormulas}");
            builder.AppendLine($"Missing cached results: {MissingCachedResults}");
            builder.AppendLine($"Dirty formulas: {DirtyFormulas}");
            builder.AppendLine();
            builder.AppendLine("| Sheet | Cell | Formula | Supported | Cached | Dirty | Reason |");
            builder.AppendLine("| --- | --- | --- | --- | --- | --- | --- |");

            foreach (ExcelFormulaCellInfo formula in Formulas) {
                builder.Append("| ");
                builder.Append(EscapeMarkdownCell(formula.SheetName));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(formula.CellReference));
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(formula.Formula));
                builder.Append(" | ");
                builder.Append(formula.IsSupportedByOfficeIMO ? "yes" : "no");
                builder.Append(" | ");
                builder.Append(formula.HasCachedValue ? "yes" : "no");
                builder.Append(" | ");
                builder.Append(formula.IsDirty ? "yes" : "no");
                builder.Append(" | ");
                builder.Append(EscapeMarkdownCell(formula.UnsupportedReason ?? string.Empty));
                builder.AppendLine(" |");
            }

            return builder.ToString();
        }

        private static string EscapeMarkdownCell(string value) {
            return value.Replace("\\", "\\\\").Replace("|", "\\|").Replace("\r", " ").Replace("\n", " ");
        }
    }

    /// <summary>
    /// Formula metadata for a single worksheet cell.
    /// </summary>
    public sealed class ExcelFormulaCellInfo {
        internal ExcelFormulaCellInfo(
            string sheetName,
            string cellReference,
            string formula,
            string? cachedValue,
            bool isDirty,
            bool isSupportedByOfficeIMO,
            string? unsupportedReason) {
            SheetName = sheetName;
            CellReference = cellReference;
            Formula = formula;
            CachedValue = cachedValue;
            IsDirty = isDirty;
            IsSupportedByOfficeIMO = isSupportedByOfficeIMO;
            UnsupportedReason = unsupportedReason;
        }

        /// <summary>Worksheet name.</summary>
        public string SheetName { get; }

        /// <summary>A1 cell reference.</summary>
        public string CellReference { get; }

        /// <summary>Formula text without forcing a leading equals sign.</summary>
        public string Formula { get; }

        /// <summary>Cached cell value, if present.</summary>
        public string? CachedValue { get; }

        /// <summary>True when a cached result is present.</summary>
        public bool HasCachedValue => CachedValue != null;

        /// <summary>True when the formula is marked for recalculation.</summary>
        public bool IsDirty { get; }

        /// <summary>True when OfficeIMO's lightweight evaluator can currently calculate this formula.</summary>
        public bool IsSupportedByOfficeIMO { get; }

        /// <summary>Reason a formula is not supported by OfficeIMO's lightweight evaluator.</summary>
        public string? UnsupportedReason { get; }
    }

    /// <summary>
    /// Describes the current lightweight formula calculation support in OfficeIMO.Excel.
    /// </summary>
    public sealed class ExcelFormulaCapabilities {
        private static readonly string[] Functions = { "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "COUNTIF", "SUMIF", "AVERAGEIF", "COUNTIFS", "SUMIFS", "AVERAGEIFS", "PRODUCT", "MEDIAN", "LARGE", "SMALL", "SUMPRODUCT", "VLOOKUP", "HLOOKUP", "XLOOKUP", "CONCAT", "TEXTJOIN", "LEFT", "RIGHT", "MID", "LEN", "TRIM", "ABS", "SIGN", "ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC", "INT", "CEILING", "FLOOR", "POWER", "SQRT", "LN", "LOG10", "EXP", "PI", "RADIANS", "DEGREES", "MOD", "DATE", "TIME", "TODAY", "NOW", "YEAR", "MONTH", "DAY", "HOUR", "MINUTE", "SECOND", "EDATE", "EOMONTH", "DAYS", "WEEKDAY", "NETWORKDAYS", "IF", "AND", "OR", "NOT", "IFERROR" };
        private static readonly string[] Operators = { "+", "-", "*", "/", ">", "<", ">=", "<=", "=", "<>" };
        private static readonly string[] OperandKinds = { "number literal", "same-sheet A1 cell reference", "same-sheet A1 range for function arguments", "cross-sheet A1 cell/range reference", "A1-backed named range reference", "simple table structured reference", "same-sheet numeric comparison for IF/AND/OR/NOT" };

        private ExcelFormulaCapabilities() {
        }

        /// <summary>Current OfficeIMO.Excel lightweight formula capability model.</summary>
        public static ExcelFormulaCapabilities Current { get; } = new ExcelFormulaCapabilities();

        /// <summary>Supported aggregate functions.</summary>
        public IReadOnlyList<string> SupportedFunctions => Functions;

        /// <summary>Supported binary arithmetic operators.</summary>
        public IReadOnlyList<string> SupportedOperators => Operators;

        /// <summary>Supported operand kinds.</summary>
        public IReadOnlyList<string> SupportedOperandKinds => OperandKinds;

        /// <summary>Maximum formula length accepted by the lightweight evaluator.</summary>
        public int MaxFormulaLength => 8192;

        /// <summary>Short human-readable summary of the current evaluator scope.</summary>
        public string Summary => "Supports simple same-sheet arithmetic plus SUM/AVERAGE/MIN/MAX/COUNT/COUNTA/COUNTIF/SUMIF/AVERAGEIF/COUNTIFS/SUMIFS/AVERAGEIFS/PRODUCT/MEDIAN/LARGE/SMALL/SUMPRODUCT, exact-match VLOOKUP/HLOOKUP/XLOOKUP returning numeric or text values, CONCAT/TEXTJOIN/LEFT/RIGHT/MID/LEN/TRIM text helpers, ABS/SIGN/ROUND/ROUNDUP/ROUNDDOWN/TRUNC/INT/CEILING/FLOOR/POWER/SQRT/LN/LOG10/EXP/PI/RADIANS/DEGREES/MOD, DATE/TIME/TODAY/NOW/YEAR/MONTH/DAY/HOUR/MINUTE/SECOND/EDATE/EOMONTH/DAYS/WEEKDAY/NETWORKDAYS, numeric IF/AND/OR/NOT comparisons, and numeric IFERROR fallbacks over numbers, text literals, A1 cells, A1 ranges, A1-backed named ranges, simple table structured references, cross-sheet references, and nested formulas.";
    }
}
