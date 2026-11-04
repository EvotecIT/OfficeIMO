using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private const int MaxFormulaDefinedNameExpansionDepth = 256;

        private bool TryResolveDefinedNameConstant(string token, out FormulaArgumentValue value) {
            value = default;
            if (!TrySplitQualifiedReference(token, out _, out string name) || !IsFormulaDefinedNameToken(name)) return false;
            var catalog = GetFormulaDefinedNameCatalog();
            var active = (_formulaEvaluationGuardState ??= new FormulaEvaluationGuardState()).ActiveDefinedNames;
            var entered = new List<string>();
            ExcelSheet resolutionSheet = this;
            try {
                for (int expansion = 0; expansion < MaxFormulaDefinedNameExpansionDepth; expansion++) {
                    if (!resolutionSheet.TryGetDefinedNameReference(catalog, token, out ExcelSheet defaultSheet,
                        out string reference, out string identity)
                        || active.Count >= MaxFormulaDefinedNameExpansionDepth || !active.Add(identity)) return false;
                    entered.Add(identity);
                    if (double.TryParse(reference, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                        value = new FormulaArgumentValue(number, reference);
                        return true;
                    }
                    if (reference.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || reference.Equals("FALSE", StringComparison.OrdinalIgnoreCase)) {
                        bool flag = reference.Equals("TRUE", StringComparison.OrdinalIgnoreCase);
                        value = new FormulaArgumentValue(flag ? 1d : 0d, flag ? "1" : "0", isBoolean: true);
                        return true;
                    }
                    if (ExcelFormulaExpressionParser.TryParseTextLiteral(reference, out string text)) {
                        value = new FormulaArgumentValue(null, text);
                        return true;
                    }
                    if (TryParseDefinedNameError(reference, catalog, out string error)) {
                        value = FormulaArgumentValue.Error(error);
                        return true;
                    }
                    if (!TrySplitQualifiedReference(reference, out _, out string alias) || !IsFormulaDefinedNameToken(alias))
                        return defaultSheet.TryEvaluateFormulaValue(reference, out value) && !value.IsUnresolvedFormula;
                    resolutionSheet = defaultSheet;
                    token = reference;
                }
                return false;
            } finally {
                foreach (string identity in entered) active.Remove(identity);
            }
        }

        private FormulaDefinedNameResolutionCatalog GetFormulaDefinedNameCatalog() =>
            _formulaCalculationContext is { IsDisposed: false } context
                ? context.GetDefinedNames(this, _excelDocument.CaptureFormulaInputMutationVersion())
                : new FormulaDefinedNameResolutionCatalog(this);

        private bool TryResolveDefinedNameArray(string token, int depth, out FormulaArrayValue array,
            bool preserveScalarValue = false) {
            array = null!;
            if (!TrySplitQualifiedReference(token, out _, out string name) || !IsFormulaDefinedNameToken(name)
                || !TryGetDefinedNameReference(GetFormulaDefinedNameCatalog(), token, out ExcelSheet sheet,
                    out string expression, out string identity)) return false;
            var active = (_formulaEvaluationGuardState ??= new FormulaEvaluationGuardState()).ActiveDefinedNames;
            if (active.Count >= MaxFormulaDefinedNameExpansionDepth || !active.Add(identity)) return false;
            try { return sheet.TryEvaluateArrayValue(expression, depth + 1, out array, preserveScalarValue); }
            finally { active.Remove(identity); }
        }

        private static bool TryParseDefinedNameError(string reference, FormulaDefinedNameResolutionCatalog catalog, out string error) {
            if (TryParseFormulaErrorLiteral(reference, out error)) return true;
            // Excel may save a broken name as a sheet-qualified #REF! value.
            int separator = reference.Length > 1 ? reference.LastIndexOf('!', reference.Length - 2) : -1;
            if (separator <= 0 || !catalog.TryGetSheet(NormalizeFormulaSheetName(reference.Substring(0, separator)), out _, out _)) return false;
            return TryParseFormulaErrorLiteral(reference.Substring(separator + 1), out error);
        }

        internal sealed class FormulaDefinedNameResolutionCatalog {
            private readonly List<Sheet> _sheets;
            private readonly HashSet<int> _validSheetIndexes = new HashSet<int>();
            private readonly Dictionary<string, int> _sheetIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, DefinedName> _definedNames = new Dictionary<string, DefinedName>(StringComparer.OrdinalIgnoreCase);

            internal FormulaDefinedNameResolutionCatalog(ExcelSheet owner) {
                _sheets = owner.WorkbookRoot.Sheets?.Elements<Sheet>().ToList() ?? new List<Sheet>();
                var worksheetRelationshipIds = new HashSet<string>(
                    owner.WorkbookPartRoot.Parts
                        .Where(pair => pair.OpenXmlPart is DocumentFormat.OpenXml.Packaging.WorksheetPart)
                        .Select(pair => pair.RelationshipId),
                    StringComparer.Ordinal);
                for (int index = 0; index < _sheets.Count; index++) {
                    string? sheetName = _sheets[index].Name?.Value;
                    string? relationshipId = _sheets[index].Id?.Value;
                    bool validWorksheet = !string.IsNullOrWhiteSpace(relationshipId)
                        && worksheetRelationshipIds.Contains(relationshipId!);
                    if (validWorksheet) _validSheetIndexes.Add(index);
                    if (validWorksheet && !string.IsNullOrWhiteSpace(sheetName) && !_sheetIndexes.ContainsKey(sheetName!)) {
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
                if (index < (uint)_sheets.Count && _validSheetIndexes.Contains((int)index)) {
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

            if (!TrySplitQualifiedReference(token, out _, out string name) || !IsFormulaDefinedNameToken(name)) return false;
            var catalog = GetFormulaDefinedNameCatalog();
            var active = (_formulaEvaluationGuardState ??= new FormulaEvaluationGuardState()).ActiveDefinedNames;
            var entered = new List<string>();
            ExcelSheet resolutionSheet = this;
            string currentToken = token;
            try {
                for (int expansion = 0; expansion < MaxFormulaDefinedNameExpansionDepth; expansion++) {
                    if (!resolutionSheet.TryGetDefinedNameReference(
                        catalog,
                        currentToken,
                        out ExcelSheet defaultSheet,
                        out string reference,
                        out string identity)) {
                        return false;
                    }

                    if (active.Count >= MaxFormulaDefinedNameExpansionDepth || !active.Add(identity)) {
                        return false;
                    }
                    entered.Add(identity);

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
            } finally {
                foreach (string identity in entered) active.Remove(identity);
            }
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
            if (ExcelFormulaExpressionParser.TryParseSupportedFunctionCall(reference, out ExcelFormulaFunctionCallSyntax? call)
                && (call!.Name.Equals("INDEX", StringComparison.OrdinalIgnoreCase)
                    || call.Name.Equals("OFFSET", StringComparison.OrdinalIgnoreCase)))
                return defaultSheet.TryResolveFormulaRangeReference(reference, currentRow, out sheet, out r1, out c1, out r2, out c2);
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

            bool isConstant = ExcelFormulaExpressionParser.TryParseTextLiteral(reference, out _)
                || TryParseDefinedNameError(reference, catalog, out _);
            if (reference.Length == 0 || (!isConstant
                && SplitFormulaArguments(reference, preserveEmpty: true).Count > 1)) {
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

        private ExcelSheet CreateDefinedNameResolutionSheet(Sheet sheetElement) {
            if (string.Equals(Name, sheetElement.Name?.Value, StringComparison.OrdinalIgnoreCase)) {
                return this;
            }

            return new ExcelSheet(_excelDocument, _spreadSheetDocument, sheetElement) {
                _formulaEvaluationCache = _formulaEvaluationCache,
                _formulaCalculationContext = _formulaCalculationContext,
                _formulaEvaluationCellReference = _formulaEvaluationCellReference,
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
