namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        /// <summary>
        /// Rewrites one formula carried by a worksheet copy or merge: cell, rule, threshold, validation, table, chart and
        /// defined-name formulas all pass through here. The formula is tokenized with <see cref="ExcelFormulaSyntaxTree"/>,
        /// so string literals are never changed, tables are renamed only in structured references and bare table names,
        /// and sheet qualifiers and external-workbook indexes are rewritten only outside string literals.
        /// </summary>
        private static string RewriteCopiedFormula(
            string? formula,
            IReadOnlyDictionary<string, string>? sheetNameMap,
            IReadOnlyDictionary<string, string>? tableNameMap = null,
            IReadOnlyDictionary<int, int>? externalReferenceMap = null) {
            if (string.IsNullOrEmpty(formula)) {
                return formula ?? string.Empty;
            }

            string text = formula!;
            if (tableNameMap?.Count > 0) {
                text = ExcelFormulaSyntaxTree.Parse(text).RewriteTableNames(name =>
                    tableNameMap.TryGetValue(name, out string? replacement) ? replacement : name);
            }

            bool rewriteSheets = sheetNameMap?.Count > 0;
            bool rewriteExternal = externalReferenceMap?.Count > 0;
            if (!rewriteSheets && !rewriteExternal) {
                return text;
            }

            // Sheet qualifiers, including 3-D ranges such as 'Jan:Mar'!A1 that span several tokens, and external
            // workbook indexes are rewritten in each run of formula text between string literals; literals stay as typed.
            return ExcelFormulaReferenceRewriter.RewriteOutsideStrings(text, segment => {
                if (rewriteSheets) {
                    segment = ReplaceSheetNameReferences(segment, sheetNameMap!);
                }

                return rewriteExternal ? RewriteExternalWorkbookIndexesInText(segment, externalReferenceMap!) : segment;
            });
        }

        /// <summary>Returns the defined names a formula uses, with any sheet qualifier, skipping string literals.</summary>
        private static IEnumerable<(string Name, string? SheetName)> ReadDefinedNameTokens(string formula) {
            int position = 0;
            foreach (ExcelFormulaSyntaxNode node in ExcelFormulaSyntaxTree.Parse(formula).Nodes) {
                int start = position;
                position += node.Text.Length;
                // In [1]!Name the workbook qualifier is tokenized apart from the name; that name lives in the
                // linked workbook. [0]! is this workbook, as charts write it.
                if (start >= 2 && formula[start - 1] == '!' && formula[start - 2] == ']'
                    && !(start >= 4 && formula[start - 3] == '0' && formula[start - 4] == '[')) {
                    continue;
                }

                string? authored = node switch {
                    ExcelFormulaNameSyntax name => name.Name,
                    // A LAMBDA stored as a defined name is called like a function. The tokenizer leaves a sheet
                    // qualifier such as Functions!Twice( outside the function token, so it is read back here.
                    ExcelFormulaFunctionSyntax function => ReadFunctionQualifier(formula, start) is string sheet
                        ? QuoteQualifier(sheet) + "!" + function.Name
                        : function.Name,
                    _ => null
                };
                if (string.IsNullOrEmpty(authored)) {
                    continue;
                }

                int bang = LastIndexOfUnquoted(authored!, '!');
                if (bang < 0) {
                    yield return (authored!, null);
                    continue;
                }

                string qualifier = authored!.Substring(0, bang);
                string localName = authored.Substring(bang + 1);
                if (qualifier.Length >= 2 && qualifier[0] == '\'' && qualifier[qualifier.Length - 1] == '\'') {
                    qualifier = qualifier.Substring(1, qualifier.Length - 2).Replace("''", "'");
                }

                // A name in another workbook belongs to that workbook, not to the copied sheet.
                if (qualifier.StartsWith("[", StringComparison.Ordinal) || localName.Length == 0) {
                    continue;
                }

                yield return (localName, qualifier);
            }
        }

        /// <summary>Reports whether a formula refers to another workbook outside its string literals.</summary>
        private static bool ContainsExternalWorkbookReference(string? formula) {
            if (string.IsNullOrEmpty(formula)) {
                return false;
            }

            // Scan whole runs between string literals: a qualifier such as [1]!Name spans more than one token.
            bool found = false;
            ExcelFormulaReferenceRewriter.RewriteOutsideStrings(formula!, segment => {
                for (int index = 0; !found && index < segment.Length; index++) {
                    found = segment[index] == '['
                        && TryReadExternalWorkbookIndex(segment, index, out _, out int closeIndex)
                        && LooksLikeExternalWorkbookReference(segment, index, closeIndex);
                }

                return segment;
            });
            return found;
        }

        /// <summary>Reads the sheet qualifier written directly before a function token, as in 'My Sheet'!Twice(.</summary>
        private static string? ReadFunctionQualifier(string formula, int functionStart) {
            int bang = functionStart - 1;
            if (bang < 1 || formula[bang] != '!') {
                return null;
            }

            int end = bang - 1;
            if (formula[end] == '\'') {
                for (int index = end - 1; index >= 0; index--) {
                    if (formula[index] != '\'') {
                        continue;
                    }

                    if (index > 0 && formula[index - 1] == '\'') {
                        index--;
                        continue;
                    }

                    return formula.Substring(index + 1, end - index - 1).Replace("''", "'");
                }

                return null;
            }

            int begin = end;
            while (begin >= 0 && (char.IsLetterOrDigit(formula[begin]) || formula[begin] == '_' || formula[begin] == '.')) {
                begin--;
            }

            if (begin == end) {
                return null;
            }

            // [1]Sheet!Fn( belongs to a linked workbook; report it with its bracket so callers skip it.
            return begin >= 0 && formula[begin] == ']' ? "[" + formula.Substring(begin + 1, end - begin) : formula.Substring(begin + 1, end - begin);
        }

        private static string QuoteQualifier(string sheetName) => "'" + sheetName.Replace("'", "''") + "'";

        private static int LastIndexOfUnquoted(string text, char value) {
            bool inQuote = false;
            int found = -1;
            for (int index = 0; index < text.Length; index++) {
                char current = text[index];
                if (current == '\'') {
                    inQuote = !inQuote;
                } else if (current == value && !inQuote) {
                    found = index;
                }
            }

            return found;
        }
    }
}
