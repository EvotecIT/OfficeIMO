using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int MaxFormulaDefinedNameExpansionDepth = 256;

        private sealed class FormulaDefinedNameResolutionCatalog {
            private readonly List<Sheet> _sheets;
            private readonly Dictionary<string, int> _sheetIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, DefinedName> _definedNames = new Dictionary<string, DefinedName>(StringComparer.OrdinalIgnoreCase);

            internal FormulaDefinedNameResolutionCatalog(ExcelSheet owner) {
                _sheets = owner.WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
                for (int index = 0; index < _sheets.Count; index++) {
                    string? sheetName = _sheets[index].Name?.Value;
                    if (!string.IsNullOrWhiteSpace(sheetName) && !_sheetIndexes.ContainsKey(sheetName!)) {
                        _sheetIndexes.Add(sheetName!, index);
                    }
                }

                foreach (DefinedName definedName in owner.WorkbookRoot.DefinedNames?.Elements<DefinedName>() ?? Enumerable.Empty<DefinedName>()) {
                    string? name = definedName.Name?.Value;
                    if (!IsFormulaDefinedNameToken(name ?? string.Empty) || IsBuiltInFormulaDefinedName(name)) {
                        continue;
                    }

                    string key = CreateDefinedNameIdentity(definedName.LocalSheetId?.Value, name!);
                    if (!_definedNames.ContainsKey(key)) {
                        _definedNames.Add(key, definedName);
                    }
                }
            }

            internal bool TryGetSheet(string name, out int index, out Sheet sheet) {
                if (_sheetIndexes.TryGetValue(name, out index)) {
                    sheet = _sheets[index];
                    return true;
                }

                sheet = null!;
                return false;
            }

            internal bool TryGetSheet(uint index, out Sheet sheet) {
                if (index < (uint)_sheets.Count) {
                    sheet = _sheets[(int)index];
                    return true;
                }

                sheet = null!;
                return false;
            }

            internal bool TryGetDefinedName(
                int? localSheetIndex,
                string name,
                bool allowGlobal,
                out DefinedName definedName,
                out string identity) {
                if (localSheetIndex.HasValue) {
                    identity = CreateDefinedNameIdentity((uint)localSheetIndex.Value, name);
                    if (_definedNames.TryGetValue(identity, out definedName!)) {
                        return true;
                    }
                }

                if (allowGlobal) {
                    identity = CreateDefinedNameIdentity(null, name);
                    if (_definedNames.TryGetValue(identity, out definedName!)) {
                        return true;
                    }
                }

                definedName = null!;
                identity = string.Empty;
                return false;
            }

            private static string CreateDefinedNameIdentity(uint? localSheetIndex, string name) {
                string scope = localSheetIndex.HasValue
                    ? localSheetIndex.Value.ToString(CultureInfo.InvariantCulture)
                    : "global";
                return scope + ":" + name;
            }
        }

        private bool TryResolveDefinedNameRange(
            string token,
            int? currentRow,
            out ExcelSheet sheet,
            out int r1,
            out int c1,
            out int r2,
            out int c2) {
            sheet = this;
            r1 = 0;
            c1 = 0;
            r2 = 0;
            c2 = 0;

            var visitedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var catalog = new FormulaDefinedNameResolutionCatalog(this);
            ExcelSheet resolutionSheet = this;
            string currentToken = token;
            for (int expansion = 0; expansion < MaxFormulaDefinedNameExpansionDepth; expansion++) {
                if (!resolutionSheet.TryGetDefinedNameReference(
                    catalog,
                    currentToken,
                    out ExcelSheet defaultSheet,
                    out string reference,
                    out string identity)) {
                    return false;
                }

                if (!visitedNames.Add(identity)) {
                    return false;
                }

                if (resolutionSheet.TryResolveDefinedNameReferenceTarget(
                    reference,
                    defaultSheet,
                    currentRow,
                    out ExcelSheet resolvedSheet,
                    out int resolvedR1,
                    out int resolvedC1,
                    out int resolvedR2,
                    out int resolvedC2)) {
                    sheet = resolvedSheet;
                    r1 = resolvedR1;
                    c1 = resolvedC1;
                    r2 = resolvedR2;
                    c2 = resolvedC2;
                    return true;
                }

                resolutionSheet = defaultSheet;
                currentToken = reference;
            }

            return false;
        }

        private bool TryResolveDefinedNameReferenceTarget(
            string reference,
            ExcelSheet defaultSheet,
            int? currentRow,
            out ExcelSheet sheet,
            out int r1,
            out int c1,
            out int r2,
            out int c2) {
            if (TryParseQualifiedFormulaRange(reference, defaultSheet, out sheet, out r1, out c1, out r2, out c2)) {
                return true;
            }

            if (TryParseQualifiedFormulaCellReference(reference, defaultSheet, out sheet, out r1, out c1)) {
                r2 = r1;
                c2 = c1;
                return true;
            }

            if (TryParseQualifiedFormulaWholeRange(
                reference,
                defaultSheet,
                out sheet,
                out r1,
                out c1,
                out r2,
                out c2,
                out _)) {
                return true;
            }

            return TryResolveTableReferenceRange(reference, currentRow, out sheet, out r1, out c1, out r2, out c2);
        }

        private bool TryGetDefinedNameReference(
            FormulaDefinedNameResolutionCatalog catalog,
            string token,
            out ExcelSheet defaultSheet,
            out string reference,
            out string identity) {
            defaultSheet = this;
            reference = string.Empty;
            identity = string.Empty;
            if (!TrySplitQualifiedReference(token, out string? sheetName, out string name)
                || !IsFormulaDefinedNameToken(name)) {
                return false;
            }

            int? localSheetIndex = null;
            if (sheetName != null) {
                if (!catalog.TryGetSheet(sheetName, out int index, out Sheet sheetElement)) {
                    return false;
                }

                localSheetIndex = index;
                defaultSheet = CreateDefinedNameResolutionSheet(sheetElement);
            } else {
                if (catalog.TryGetSheet(Name, out int index, out _)) {
                    localSheetIndex = index;
                }
            }

            if (!catalog.TryGetDefinedName(
                localSheetIndex,
                name,
                allowGlobal: sheetName == null,
                out DefinedName definedName,
                out identity)) {
                return false;
            }

            reference = (definedName.Text ?? string.Empty).Trim();
            if (reference.StartsWith("=", StringComparison.Ordinal)) {
                reference = reference.Substring(1).Trim();
            }

            if (reference.Length == 0
                || ContainsTopLevelFormulaComma(reference)
                || reference.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0) {
                return false;
            }

            if (definedName.LocalSheetId?.Value is uint scopedIndex
                && catalog.TryGetSheet(scopedIndex, out Sheet scopedSheet)
                && scopedSheet.Name?.Value is string scopedSheetName
                && !string.Equals(scopedSheetName, Name, StringComparison.OrdinalIgnoreCase)) {
                defaultSheet = CreateDefinedNameResolutionSheet(scopedSheet);
            }

            return true;
        }

        private static bool ContainsTopLevelFormulaComma(string reference) {
            bool inQuotedQualifier = false;
            int bracketDepth = 0;
            for (int index = 0; index < reference.Length; index++) {
                char character = reference[index];
                if (character == '\'') {
                    if (inQuotedQualifier
                        && index + 1 < reference.Length
                        && reference[index + 1] == '\'') {
                        index++;
                        continue;
                    }

                    inQuotedQualifier = !inQuotedQualifier;
                } else if (!inQuotedQualifier && character == '[') {
                    bracketDepth++;
                } else if (!inQuotedQualifier && character == ']' && bracketDepth > 0) {
                    bracketDepth--;
                } else if (character == ',' && !inQuotedQualifier && bracketDepth == 0) {
                    return true;
                }
            }

            return false;
        }

        private ExcelSheet CreateDefinedNameResolutionSheet(Sheet sheetElement) {
            if (string.Equals(Name, sheetElement.Name?.Value, StringComparison.OrdinalIgnoreCase)) {
                return this;
            }

            return new ExcelSheet(_excelDocument, _spreadSheetDocument, sheetElement) {
                _formulaEvaluationCache = _formulaEvaluationCache,
                _formulaEvaluationDepthCache = _formulaEvaluationDepthCache,
                _formulaEvaluationStack = _formulaEvaluationStack,
                _formulaEvaluationDepthFrames = _formulaEvaluationDepthFrames,
                _formulaEvaluationGuardState = _formulaEvaluationGuardState
            };
        }

        private static bool IsFormulaDefinedNameToken(string token) {
            if (string.IsNullOrWhiteSpace(token)
                || string.Equals(token, "TRUE", StringComparison.OrdinalIgnoreCase)
                || string.Equals(token, "FALSE", StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            char first = token[0];
            if (!char.IsLetter(first) && first != '_' && first != '\\') {
                return false;
            }

            foreach (char character in token) {
                if (!char.IsLetterOrDigit(character)
                    && character != '_'
                    && character != '.'
                    && character != '\\') {
                    return false;
                }
            }

            return true;
        }

        private static bool IsBuiltInFormulaDefinedName(string? name) {
            return !string.IsNullOrWhiteSpace(name)
                && name!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase);
        }
    }
}
