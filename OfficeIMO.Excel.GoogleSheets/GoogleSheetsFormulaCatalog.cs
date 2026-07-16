namespace OfficeIMO.Excel.GoogleSheets {
    /// <summary>
    /// Describes how one Excel formula maps to Google Sheets.
    /// </summary>
    public sealed class GoogleSheetsFormulaTranslation {
        internal GoogleSheetsFormulaTranslation(string formula, IReadOnlyList<string> functions, IReadOnlyList<string> unsupportedFunctions) {
            Formula = formula;
            Functions = functions;
            UnsupportedFunctions = unsupportedFunctions;
        }

        public string Formula { get; }
        public IReadOnlyList<string> Functions { get; }
        public IReadOnlyList<string> UnsupportedFunctions { get; }
        public bool IsSupported => UnsupportedFunctions.Count == 0;
    }

    /// <summary>
    /// Conservative Excel-to-Sheets function catalog and formula rewriter.
    /// </summary>
    public static class GoogleSheetsFormulaCatalog {
        private static readonly HashSet<string> SupportedFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "ABS", "ADDRESS", "AND", "ARRAYFORMULA", "AVERAGE", "AVERAGEIF", "AVERAGEIFS",
            "CEILING", "CHAR", "CHOOSE", "CLEAN", "COLUMN", "COLUMNS", "CONCAT", "CONCATENATE",
            "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "DATE", "DATEDIF", "DATEVALUE",
            "DAY", "DAYS", "EDATE", "EOMONTH", "ERROR.TYPE", "EVEN", "EXACT", "EXP", "FILTER",
            "FIND", "FLOOR", "HLOOKUP", "HOUR", "HYPERLINK", "IF", "IFERROR", "IFNA", "IFS",
            "IMAGE", "INDEX", "INDIRECT", "INT", "ISBLANK", "ISERROR", "ISEVEN", "ISLOGICAL",
            "ISNA", "ISNUMBER", "ISODD", "ISTEXT", "LEFT", "LEN", "LN", "LOG", "LOOKUP", "LOWER",
            "MATCH", "MAX", "MEDIAN", "MID", "MIN", "MINUTE", "MOD", "MONTH", "MROUND", "NA",
            "NETWORKDAYS", "NOT", "NOW", "ODD", "OFFSET", "OR", "POWER", "PROPER", "QUERY",
            "RAND", "RANDBETWEEN", "REGEXEXTRACT", "REGEXMATCH", "REGEXREPLACE", "REPLACE", "RIGHT",
            "ROUND", "ROUNDDOWN", "ROUNDUP", "ROW", "ROWS", "SEARCH", "SECOND", "SEQUENCE", "SORT",
            "SQRT", "STDEV", "STDEV.P", "STDEV.S", "SUBSTITUTE", "SUBTOTAL", "SUM", "SUMIF", "SUMIFS",
            "SUMPRODUCT", "TEXT", "TEXTJOIN", "TIME", "TODAY", "TRANSPOSE", "TRIM", "UNIQUE", "UPPER",
            "VALUE", "VLOOKUP", "WEEKDAY", "WEEKNUM", "WORKDAY", "XLOOKUP", "XMATCH", "YEAR"
        };

        private static readonly Dictionary<string, string> BuiltInMappings = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
            ["_XLFN.CONCAT"] = "CONCAT",
            ["_XLFN.IFS"] = "IFS",
            ["_XLFN.TEXTJOIN"] = "TEXTJOIN",
            ["_XLFN.XLOOKUP"] = "XLOOKUP",
            ["_XLFN.XMATCH"] = "XMATCH",
            ["_XLFN.FILTER"] = "FILTER",
            ["_XLFN.SORT"] = "SORT",
            ["_XLFN.UNIQUE"] = "UNIQUE",
            ["_XLFN.SEQUENCE"] = "SEQUENCE",
        };

        private static readonly string[] UnsupportedPrefixes = {
            "CUBE", "_XLWS.", "RTD", "CALL", "REGISTER.", "WEBSERVICE"
        };

        public static GoogleSheetsFormulaTranslation Translate(string formula, GoogleSheetsFormulaOptions? options = null) {
            if (string.IsNullOrWhiteSpace(formula)) {
                return new GoogleSheetsFormulaTranslation("=", Array.Empty<string>(), Array.Empty<string>());
            }

            GoogleSheetsFormulaOptions effective = options ?? new GoogleSheetsFormulaOptions();
            string normalized = formula[0] == '=' ? formula : "=" + formula;
            IReadOnlyList<FormulaToken> tokens = FindFunctionTokens(normalized);
            var functions = new List<string>();
            var unsupported = new List<string>();
            var replacements = new List<(int Start, int Length, string Value)>();

            foreach (FormulaToken token in tokens) {
                string original = token.Name;
                string mapped = ResolveMapping(original, effective);
                functions.Add(mapped);
                if (!string.Equals(original, mapped, StringComparison.Ordinal)) {
                    replacements.Add((token.Start, token.Length, mapped));
                }

                bool explicitlyUnsupported = UnsupportedPrefixes.Any(prefix => mapped.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
                bool unknown = effective.TreatUnknownFunctionsAsUnsupported && !SupportedFunctions.Contains(mapped);
                if ((explicitlyUnsupported || unknown) && !unsupported.Contains(mapped, StringComparer.OrdinalIgnoreCase)) {
                    unsupported.Add(mapped);
                }
            }

            for (int index = replacements.Count - 1; index >= 0; index--) {
                var replacement = replacements[index];
                normalized = normalized.Remove(replacement.Start, replacement.Length).Insert(replacement.Start, replacement.Value);
            }

            return new GoogleSheetsFormulaTranslation(
                normalized,
                functions.Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray(),
                unsupported.OrderBy(name => name, StringComparer.OrdinalIgnoreCase).ToArray());
        }

        private static string ResolveMapping(string function, GoogleSheetsFormulaOptions options) {
            if (options.FunctionMappings.TryGetValue(function, out string? custom) && !string.IsNullOrWhiteSpace(custom)) {
                return custom.Trim().ToUpperInvariant();
            }
            return BuiltInMappings.TryGetValue(function, out string? builtIn) ? builtIn : function.ToUpperInvariant();
        }

        private static IReadOnlyList<FormulaToken> FindFunctionTokens(string formula) {
            var tokens = new List<FormulaToken>();
            bool inDoubleQuote = false;
            bool inSingleQuote = false;
            for (int index = 0; index < formula.Length; index++) {
                char current = formula[index];
                if (current == '"' && !inSingleQuote) {
                    if (inDoubleQuote && index + 1 < formula.Length && formula[index + 1] == '"') {
                        index++;
                        continue;
                    }
                    inDoubleQuote = !inDoubleQuote;
                    continue;
                }
                if (current == '\'' && !inDoubleQuote) {
                    if (inSingleQuote && index + 1 < formula.Length && formula[index + 1] == '\'') {
                        index++;
                        continue;
                    }
                    inSingleQuote = !inSingleQuote;
                    continue;
                }
                if (inDoubleQuote || inSingleQuote || !IsFunctionStart(current)) continue;

                int start = index;
                while (index + 1 < formula.Length && IsFunctionPart(formula[index + 1])) index++;
                int end = index + 1;
                int cursor = end;
                while (cursor < formula.Length && char.IsWhiteSpace(formula[cursor])) cursor++;
                if (cursor < formula.Length && formula[cursor] == '(') {
                    tokens.Add(new FormulaToken(start, end - start, formula.Substring(start, end - start)));
                }
            }
            return tokens;
        }

        private static bool IsFunctionStart(char value) => char.IsLetter(value) || value == '_';
        private static bool IsFunctionPart(char value) => char.IsLetterOrDigit(value) || value == '_' || value == '.';

        private readonly struct FormulaToken {
            public FormulaToken(int start, int length, string name) {
                Start = start;
                Length = length;
                Name = name;
            }
            public int Start { get; }
            public int Length { get; }
            public string Name { get; }
        }
    }
}
