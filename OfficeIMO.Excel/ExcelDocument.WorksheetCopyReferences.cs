using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeFormula = DocumentFormat.OpenXml.Office.Excel.Formula;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private void RewriteCopiedWorksheetReferences(
            WorksheetPart worksheetPart,
            IReadOnlyDictionary<string, string> sheetNameMap,
            IReadOnlyDictionary<string, string>? tableNameMap = null,
            string? targetSheetName = null) {
            Worksheet worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            bool worksheetChanged = RewriteWorksheetSheetReferences(worksheet, sheetNameMap);
            if (tableNameMap?.Count > 0) {
                worksheetChanged |= RewriteStructuredTableReferences(worksheet, tableNameMap);
            }

            foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                Table? table = tablePart.Table;
                if (table == null) {
                    continue;
                }

                bool tableChanged = false;
                foreach (CalculatedColumnFormula formula in table.Descendants<CalculatedColumnFormula>()) {
                    tableChanged |= RewriteFormulaSheetReference(formula, sheetNameMap);
                }

                foreach (TotalsRowFormula formula in table.Descendants<TotalsRowFormula>()) {
                    tableChanged |= RewriteFormulaSheetReference(formula, sheetNameMap);
                }

                if (tableNameMap?.Count > 0) {
                    tableChanged |= RewriteStructuredTableReferences(table, tableNameMap);
                }

                if (tableChanged) {
                    table.Save();
                }
            }

            if (worksheetChanged) {
                worksheet.Save();
            }

            RewriteCopiedChartReferences(worksheetPart, sheetNameMap, tableNameMap, targetSheetName: targetSheetName);
        }

        /// <summary>Rewrites references in both standard and extended chart series.</summary>
        private static void RewriteCopiedChartReferences(WorksheetPart worksheetPart,
            IReadOnlyDictionary<string, string>? sheetNameMap = null,
            IReadOnlyDictionary<string, string>? tableNameMap = null,
            IReadOnlyDictionary<int, int>? externalReferenceMap = null,
            string? targetSheetName = null) {
            DrawingsPart? drawingsPart = worksheetPart.DrawingsPart;
            if (drawingsPart == null) return;
            foreach (OpenXmlPart part in drawingsPart.ChartParts.Cast<OpenXmlPart>().Concat(drawingsPart.ExtendedChartParts)) {
                OpenXmlPartRootElement? root = part.RootElement;
                if (root == null) continue;
                bool changed = false;
                foreach (OpenXmlLeafTextElement formula in ChartFormulaElements(root)) {
                    if (sheetNameMap != null) changed |= RewriteFormulaSheetReference(formula, sheetNameMap);
                    // Excel stores a current-workbook chart name as [0]!Name. Copied names are
                    // sheet-scoped, so bind this formula to the copied sheet's name.
                    if (targetSheetName != null && formula.Text.StartsWith("[0]!", StringComparison.Ordinal)) {
                        formula.Text = "'" + targetSheetName.Replace("'", "''") + "'!" + formula.Text.Substring(4);
                        changed = true;
                    }
                    if (tableNameMap != null) changed |= RewriteStructuredTableFormula(formula, tableNameMap);
                    if (externalReferenceMap != null) changed |= RewriteExternalWorkbookReferenceIndexFormula(formula, externalReferenceMap);
                }
                if (changed) root.Save();
            }
        }

        private static IEnumerable<OpenXmlLeafTextElement> ChartFormulaElements(OpenXmlElement root) {
            return root.Descendants<DocumentFormat.OpenXml.Drawing.Charts.Formula>().Cast<OpenXmlLeafTextElement>()
                .Concat(root.Descendants<DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing.Formula>());
        }

        private static bool RewriteStructuredTableReferences(Worksheet worksheet, IReadOnlyDictionary<string, string> tableNameMap) {
            if (!HasRenamedTables(tableNameMap)) return false;
            bool changed = false;
            foreach (CellFormula formula in worksheet.Descendants<CellFormula>())
                changed |= RewriteStructuredTableFormula(formula, tableNameMap);
            foreach (Formula formula in worksheet.Descendants<Formula>())
                changed |= RewriteStructuredTableFormula(formula, tableNameMap);
            foreach (Formula1 formula in worksheet.Descendants<Formula1>())
                changed |= RewriteStructuredTableFormula(formula, tableNameMap);
            foreach (Formula2 formula in worksheet.Descendants<Formula2>())
                changed |= RewriteStructuredTableFormula(formula, tableNameMap);
            foreach (ConditionalFormatValueObject threshold in FormulaThresholds(worksheet)) {
                string rewritten = RewriteStructuredTableReferences(threshold.Val!.Value!, tableNameMap);
                if (!string.Equals(rewritten, threshold.Val.Value, StringComparison.Ordinal)) {
                    threshold.Val = rewritten;
                    changed = true;
                }
            }
            foreach (OfficeFormula formula in worksheet.Descendants<OfficeFormula>())
                changed |= RewriteStructuredTableFormula(formula, tableNameMap);
            return changed;
        }

        private static bool RewriteStructuredTableReferences(Table table, IReadOnlyDictionary<string, string> tableNameMap) {
            if (!HasRenamedTables(tableNameMap)) return false;
            bool changed = false;
            foreach (CalculatedColumnFormula formula in table.Descendants<CalculatedColumnFormula>())
                changed |= RewriteStructuredTableFormula(formula, tableNameMap);
            foreach (TotalsRowFormula formula in table.Descendants<TotalsRowFormula>())
                changed |= RewriteStructuredTableFormula(formula, tableNameMap);
            return changed;
        }

        private static bool HasRenamedTables(IReadOnlyDictionary<string, string> tableNameMap) {
            foreach (var pair in tableNameMap) {
                if (!string.Equals(pair.Key, pair.Value, StringComparison.Ordinal)) return true;
            }
            return false;
        }

        private static bool RewriteStructuredTableFormula(OpenXmlLeafTextElement formula, IReadOnlyDictionary<string, string> tableNameMap) {
            string rewritten = RewriteStructuredTableReferences(formula.Text, tableNameMap);
            if (string.Equals(formula.Text, rewritten, StringComparison.Ordinal)) return false;
            formula.Text = rewritten;
            return true;
        }

        private void CopyReferencedDefinedNamesFromSource(
            ExcelDocument sourceDocument,
            ExcelSheet targetSheet,
            IReadOnlyDictionary<string, string> sheetNameMap,
            IReadOnlyDictionary<string, string>? tableNameMap = null,
            IReadOnlyDictionary<int, int>? externalReferenceMap = null,
            bool layoutFormulasOnly = false) {
            ushort targetSheetPosition = GetSheetPositionIndex(targetSheet);
            var sourceSheetNamesByPosition = sourceDocument.GetSheetNamesByPosition();
            string? currentSourceSheetName = sheetNameMap
                .FirstOrDefault(mapping => string.Equals(mapping.Value, targetSheet.Name, StringComparison.OrdinalIgnoreCase))
                .Key;
            if (string.IsNullOrEmpty(currentSourceSheetName)) {
                return;
            }

            ExcelSheet sourceSheet = sourceDocument.GetSheet(currentSourceSheetName!);
            var copiedSourceSheetNames = new HashSet<string>(sheetNameMap.Keys, StringComparer.OrdinalIgnoreCase);
            IReadOnlyList<DefinedName> referencedSourceNames = ResolveReferencedDefinedNamesFromSource(
                sourceDocument,
                sourceSheet,
                copiedSourceSheetNames,
                layoutFormulasOnly);
            var plannedCopies = new List<(DefinedName Clone, ushort DestinationSheetPosition, string Name)>();
            foreach (DefinedName sourceName in referencedSourceNames) {
                string name = sourceName.Name!.Value!;
                ushort destinationSheetPosition = targetSheetPosition;
                if (sourceName.LocalSheetId != null) {
                    if (!sourceSheetNamesByPosition.TryGetValue((ushort)sourceName.LocalSheetId.Value, out string? sourceNameOwner)
                        || !sheetNameMap.TryGetValue(sourceNameOwner, out string? targetNameOwner)
                        || !TryGetSheetPositionIndexByName(targetNameOwner, out destinationSheetPosition)) {
                        continue;
                    }
                }

                var clone = (DefinedName)sourceName.CloneNode(true);
                clone.LocalSheetId = destinationSheetPosition;
                clone.Name = name;
                if (!string.IsNullOrEmpty(clone.Text)) {
                    clone.Text = ReplaceSheetNameReferences(clone.Text!, sheetNameMap);
                    if (tableNameMap?.Count > 0) {
                        clone.Text = RewriteStructuredTableReferences(clone.Text!, tableNameMap);
                    }

                    if (externalReferenceMap?.Count > 0) {
                        clone.Text = RewriteExternalWorkbookReferenceIndexes(clone.Text!, externalReferenceMap);
                    }
                }

                plannedCopies.Add((clone, destinationSheetPosition, name));
            }

            if (plannedCopies.Count == 0) {
                return;
            }

            DefinedNames targetDefinedNames = WorkbookRoot.DefinedNames ??= new DefinedNames();
            foreach (var plannedCopy in plannedCopies) {
                foreach (DefinedName existing in targetDefinedNames.Elements<DefinedName>()
                    .Where(item => item.LocalSheetId != null
                        && item.LocalSheetId.Value == plannedCopy.DestinationSheetPosition
                        && string.Equals(item.Name?.Value, plannedCopy.Name, StringComparison.OrdinalIgnoreCase))
                    .ToList()) {
                    existing.Remove();
                }

                targetDefinedNames.Append(plannedCopy.Clone);
            }

            WorkbookRoot.Save();
        }

        private static void PreflightReferencedDefinedNamesFromSource(
            ExcelDocument sourceDocument,
            IReadOnlyList<ExcelSheet> sourceSheets,
            DefinedNameCopyBudget budget,
            bool layoutFormulasOnly = false) {
            var copiedSourceSheetNames = new HashSet<string>(
                sourceSheets.Select(sheet => sheet.Name),
                StringComparer.OrdinalIgnoreCase);
            foreach (ExcelSheet sourceSheet in sourceSheets) {
                IReadOnlyList<DefinedName> referencedNames = ResolveReferencedDefinedNamesFromSource(
                    sourceDocument,
                    sourceSheet,
                    copiedSourceSheetNames,
                    layoutFormulasOnly);
                foreach (DefinedName sourceName in referencedNames) {
                    budget.Consume(sourceName.Text?.Length ?? 0);
                }

                foreach (DefinedName printName in GetSourcePrintDefinedNames(sourceDocument, sourceSheet.Name)) {
                    budget.Consume(printName.Text?.Length ?? 0);
                }
            }
        }

        private static IReadOnlyList<DefinedName> ResolveReferencedDefinedNamesFromSource(
            ExcelDocument sourceDocument,
            ExcelSheet sourceSheet,
            ISet<string> copiedSourceSheetNames,
            bool layoutFormulasOnly) {
            DefinedNames? sourceDefinedNames = sourceDocument.WorkbookRoot.DefinedNames;
            if (sourceDefinedNames == null) {
                return Array.Empty<DefinedName>();
            }

            List<string> formulaTexts = (layoutFormulasOnly
                ? CollectLayoutFormulaTexts(sourceSheet.WorksheetPart)
                : CollectFormulaTexts(sourceSheet.WorksheetPart)).ToList();
            // Every route copies the print area and titles, so names they use (a dynamic print area) come along too.
            formulaTexts.AddRange(GetSourcePrintDefinedNames(sourceDocument, sourceSheet.Name).Select(name => name.Text));
            if (formulaTexts.Count == 0) {
                return Array.Empty<DefinedName>();
            }

            var sourceSheetNamesByPosition = sourceDocument.GetSheetNamesByPosition();
            List<DefinedName> sourceNames = sourceDefinedNames.Elements<DefinedName>()
                .Where(name => !string.IsNullOrWhiteSpace(name.Name?.Value)
                    && !name.Name!.Value!.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase))
                .ToList();
            var sourceNamesByName = sourceNames
                .GroupBy(name => name.Name!.Value!, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.OrdinalIgnoreCase);
            var pendingNames = new Queue<DefinedNameReference>();
            var queuedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string formulaText in formulaTexts) {
                EnqueueReferencedDefinedNames(
                    formulaText,
                    sourceSheet.Name,
                    sourceNamesByName,
                    queuedNames,
                    pendingNames);
            }

            var resolvedNames = new List<DefinedName>();
            var copiedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            while (pendingNames.Count > 0) {
                DefinedNameReference reference = pendingNames.Dequeue();
                List<DefinedName> candidates = sourceNamesByName[reference.Name];
                bool hasLocalNameInContext = candidates.Any(candidate =>
                    candidate.LocalSheetId != null
                    && sourceSheetNamesByPosition.TryGetValue((ushort)candidate.LocalSheetId.Value, out string? candidateOwner)
                    && string.Equals(candidateOwner, reference.ContextSheetName, StringComparison.OrdinalIgnoreCase));
                foreach (DefinedName sourceName in candidates) {
                    string name = sourceName.Name!.Value!;
                    string dependencyContext = reference.ContextSheetName;
                    if (sourceName.LocalSheetId == null) {
                        if (reference.ExplicitSheetName != null || hasLocalNameInContext) {
                            continue;
                        }
                    } else {
                        if (!sourceSheetNamesByPosition.TryGetValue((ushort)sourceName.LocalSheetId.Value, out string? localOwner)) {
                            continue;
                        }

                        string requiredOwner = reference.ExplicitSheetName ?? reference.ContextSheetName;
                        if (!string.Equals(localOwner, requiredOwner, StringComparison.OrdinalIgnoreCase)) {
                            continue;
                        }

                        dependencyContext = localOwner;
                    }

                    string copyKey = name + "|" + (sourceName.LocalSheetId?.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty);
                    if (!copiedNames.Add(copyKey)) {
                        continue;
                    }

                    if (sourceName.LocalSheetId != null
                        && (!sourceSheetNamesByPosition.TryGetValue((ushort)sourceName.LocalSheetId.Value, out string? copiedOwner)
                            || !copiedSourceSheetNames.Contains(copiedOwner))) {
                        continue;
                    }

                    resolvedNames.Add(sourceName);
                    if (!string.IsNullOrEmpty(sourceName.Text)) {
                        EnqueueReferencedDefinedNames(
                            sourceName.Text!,
                            dependencyContext,
                            sourceNamesByName,
                            queuedNames,
                            pendingNames);
                    }
                }
            }

            return resolvedNames;
        }

        private static void EnqueueReferencedDefinedNames(
            string formula,
            string contextSheetName,
            IReadOnlyDictionary<string, List<DefinedName>> sourceNamesByName,
            ISet<string> queuedNames,
            Queue<DefinedNameReference> pendingNames) {
            int tokenStart = -1;
            for (int index = 0; index <= formula.Length; index++) {
                bool isTokenCharacter = index < formula.Length && IsDefinedNameCharacter(formula[index]);
                if (isTokenCharacter) {
                    if (tokenStart < 0) {
                        tokenStart = index;
                    }

                    continue;
                }

                if (tokenStart < 0) {
                    continue;
                }

                string token = formula.Substring(tokenStart, index - tokenStart);
                tokenStart = -1;
                if (!sourceNamesByName.ContainsKey(token)) {
                    continue;
                }

                string? explicitSheetName = TryReadSheetQualifier(formula, index - token.Length);
                var reference = new DefinedNameReference(token, contextSheetName, explicitSheetName);
                if (queuedNames.Add(reference.Key)) {
                    pendingNames.Enqueue(reference);
                }
            }
        }

        private static string? TryReadSheetQualifier(string formula, int tokenStart) {
            int bangIndex = tokenStart - 1;
            if (bangIndex < 0 || formula[bangIndex] != '!') {
                return null;
            }

            int qualifierEnd = bangIndex - 1;
            if (qualifierEnd < 0) {
                return null;
            }

            if (formula[qualifierEnd] == '\'') {
                for (int index = qualifierEnd - 1; index >= 0; index--) {
                    if (formula[index] != '\'') {
                        continue;
                    }

                    if (index > 0 && formula[index - 1] == '\'') {
                        index--;
                        continue;
                    }

                    return formula.Substring(index + 1, qualifierEnd - index - 1).Replace("''", "'");
                }

                return null;
            }

            int qualifierStart = qualifierEnd;
            while (qualifierStart >= 0 && IsDefinedNameCharacter(formula[qualifierStart])) {
                qualifierStart--;
            }

            return qualifierStart == qualifierEnd
                ? null
                : formula.Substring(qualifierStart + 1, qualifierEnd - qualifierStart);
        }

        private readonly struct DefinedNameReference {
            internal DefinedNameReference(string name, string contextSheetName, string? explicitSheetName) {
                Name = name;
                ContextSheetName = contextSheetName;
                ExplicitSheetName = explicitSheetName;
            }

            internal string Name { get; }

            internal string ContextSheetName { get; }

            internal string? ExplicitSheetName { get; }

            internal string Key => ContextSheetName + "|" + (ExplicitSheetName ?? string.Empty) + "|" + Name;
        }

        private sealed class DefinedNameCopyBudget {
            private readonly int _maximumNames;
            private readonly int _maximumCharacters;
            private int _copiedNames;
            private long _copiedCharacters;

            internal DefinedNameCopyBudget(int maximumNames, int maximumCharacters) {
                _maximumNames = maximumNames;
                _maximumCharacters = maximumCharacters;
            }

            internal void Consume(int characters) {
                if (_copiedNames >= _maximumNames) {
                    throw new InvalidOperationException(
                        $"Worksheet copy exceeds the configured defined-name limit of {_maximumNames}.");
                }

                long totalCharacters = _copiedCharacters + characters;
                if (totalCharacters > _maximumCharacters) {
                    throw new InvalidOperationException(
                        $"Worksheet copy exceeds the configured defined-name character limit of {_maximumCharacters}.");
                }

                _copiedNames++;
                _copiedCharacters = totalCharacters;
            }
        }

        private Dictionary<ushort, string> GetSheetNamesByPosition() {
            var names = new Dictionary<ushort, string>();
            ushort position = 0;
            foreach (Sheet sheet in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                string? name = sheet.Name?.Value;
                if (!string.IsNullOrEmpty(name)) {
                    names[position] = name!;
                }

                position++;
            }

            return names;
        }

        private bool TryGetSheetPositionIndexByName(string sheetName, out ushort position) {
            position = 0;
            foreach (Sheet sheet in WorkbookRoot.Sheets?.Elements<Sheet>() ?? Enumerable.Empty<Sheet>()) {
                if (string.Equals(sheet.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }

                position++;
            }

            return false;
        }

        /// <summary>
        /// Collects the formulas a values copy keeps: conditional-format and data-validation rules. Cell and table
        /// formulas are not copied in values mode, so the names they use are not needed.
        /// </summary>
        /// <remarks>A worksheet that is not loaded is streamed, so a values copy never loads the source cells.</remarks>
        private static List<string> CollectLayoutFormulaTexts(WorksheetPart worksheetPart) {
            var formulas = new List<string>();
            if (worksheetPart.IsRootElementLoaded) {
                Worksheet worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
                foreach (OpenXmlElement container in worksheet.Elements<ConditionalFormatting>().Cast<OpenXmlElement>().Concat(worksheet.Elements<DataValidations>())) {
                    AddLayoutFormulaTexts(container, formulas);
                }

                return formulas;
            }

            using OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
            bool hasCurrent = reader.Read();
            while (hasCurrent) {
                if (reader.IsStartElement && reader.Depth == 1
                    && (reader.ElementType == typeof(ConditionalFormatting) || reader.ElementType == typeof(DataValidations))) {
                    AddLayoutFormulaTexts(reader.LoadCurrentElement()!, formulas);
                }

                hasCurrent = reader.Read();
            }

            return formulas;
        }

        private static void AddLayoutFormulaTexts(OpenXmlElement container, List<string> formulas) {
            foreach (OpenXmlLeafTextElement formula in container.Descendants<OpenXmlLeafTextElement>()) {
                if ((formula is Formula || formula is Formula1 || formula is Formula2) && !string.IsNullOrEmpty(formula.Text)) {
                    formulas.Add(formula.Text);
                }
            }

            formulas.AddRange(FormulaThresholds(container).Select(threshold => threshold.Val!.Value!));
        }

        /// <summary>Color-scale, data-bar and icon-set thresholds whose value is a formula.</summary>
        private static IEnumerable<ConditionalFormatValueObject> FormulaThresholds(OpenXmlElement root) =>
            root.Descendants<ConditionalFormatValueObject>()
                .Where(threshold => threshold.Type?.Value == ConditionalFormatValueObjectValues.Formula
                    && !string.IsNullOrEmpty(threshold.Val?.Value));

        /// <summary>
        /// Reports whether the rules a values copy keeps, or the defined names they use, refer to another workbook.
        /// </summary>
        private static bool ValuesCopyRulesReferenceExternalWorkbooks(ExcelDocument sourceDocument, IReadOnlyList<ExcelSheet> sourceSheets) {
            if (sourceDocument.WorkbookRoot.ExternalReferences?.Elements<ExternalReference>().Any() != true) {
                return false;
            }

            var copiedSourceSheetNames = new HashSet<string>(sourceSheets.Select(sheet => sheet.Name), StringComparer.OrdinalIgnoreCase);
            foreach (ExcelSheet sourceSheet in sourceSheets) {
                if (CollectLayoutFormulaTexts(sourceSheet.WorksheetPart).Any(ContainsExternalWorkbookReference)
                    || GetSourcePrintDefinedNames(sourceDocument, sourceSheet.Name).Any(name => ContainsExternalWorkbookReference(name.Text))) {
                    return true;
                }

                foreach (DefinedName name in ResolveReferencedDefinedNamesFromSource(sourceDocument, sourceSheet, copiedSourceSheetNames, layoutFormulasOnly: true)) {
                    if (ContainsExternalWorkbookReference(name.Text)) {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Applies the package-mode rule for external workbook links to a values copy: rules that use them are refused
        /// unless the caller opts in, and are then copied with their links and remapped indexes.
        /// </summary>
        /// <returns>The source-to-target external link index map, empty when nothing needs remapping.</returns>
        private IReadOnlyDictionary<int, int> PrepareValuesCopyExternalWorkbookReferences(
            ExcelDocument sourceDocument,
            IReadOnlyList<ExcelSheet> sourceSheets,
            bool copyExternalWorkbookReferences) {
            if (ReferenceEquals(sourceDocument, this) || !ValuesCopyRulesReferenceExternalWorkbooks(sourceDocument, sourceSheets)) {
                return new Dictionary<int, int>();
            }

            if (!copyExternalWorkbookReferences) {
                throw new InvalidOperationException(
                    "Values-mode worksheet copy refuses conditional formats, data validations and defined names that reference external workbooks unless CopyExternalWorkbookReferences is enabled explicitly.");
            }

            return Locking.ExecuteWrite(EnsureLock(), () => CopyExternalWorkbookReferencesFromSource(sourceDocument));
        }

        private static bool ContainsExternalWorkbookReference(string? formula) {
            if (string.IsNullOrEmpty(formula)) {
                return false;
            }

            bool inStringLiteral = false;
            for (int index = 0; index < formula!.Length; index++) {
                char current = formula[index];
                if (current == '"') {
                    inStringLiteral = !inStringLiteral;
                } else if (!inStringLiteral
                    && current == '['
                    && TryReadExternalWorkbookIndex(formula, index, out _, out int closeIndex)
                    && LooksLikeExternalWorkbookReference(formula, index, closeIndex)) {
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<string> CollectFormulaTexts(WorksheetPart worksheetPart) {
            Worksheet worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            foreach (CellFormula formula in worksheet.Descendants<CellFormula>()) {
                if (!string.IsNullOrEmpty(formula.Text)) {
                    yield return formula.Text!;
                }
            }

            foreach (Formula formula in worksheet.Descendants<Formula>()) {
                if (!string.IsNullOrEmpty(formula.Text)) {
                    yield return formula.Text!;
                }
            }

            foreach (Formula1 formula in worksheet.Descendants<Formula1>()) {
                if (!string.IsNullOrEmpty(formula.Text)) {
                    yield return formula.Text!;
                }
            }

            foreach (Formula2 formula in worksheet.Descendants<Formula2>()) {
                if (!string.IsNullOrEmpty(formula.Text)) {
                    yield return formula.Text!;
                }
            }

            foreach (ConditionalFormatValueObject threshold in FormulaThresholds(worksheet)) {
                yield return threshold.Val!.Value!;
            }

            foreach (OfficeFormula formula in worksheet.Descendants<OfficeFormula>()) {
                if (!string.IsNullOrEmpty(formula.Text)) {
                    yield return formula.Text!;
                }
            }

            if (worksheetPart.DrawingsPart is DrawingsPart drawings) {
                foreach (OpenXmlPart part in drawings.ChartParts.Cast<OpenXmlPart>().Concat(drawings.ExtendedChartParts)) {
                    if (part.RootElement == null) continue;
                    foreach (OpenXmlLeafTextElement formula in ChartFormulaElements(part.RootElement)) {
                        if (!string.IsNullOrEmpty(formula.Text)) yield return formula.Text;
                    }
                }
            }

            foreach (TableDefinitionPart tablePart in worksheetPart.TableDefinitionParts) {
                Table? table = tablePart.Table;
                if (table == null) {
                    continue;
                }

                foreach (CalculatedColumnFormula formula in table.Descendants<CalculatedColumnFormula>()) {
                    if (!string.IsNullOrEmpty(formula.Text)) {
                        yield return formula.Text!;
                    }
                }

                foreach (TotalsRowFormula formula in table.Descendants<TotalsRowFormula>()) {
                    if (!string.IsNullOrEmpty(formula.Text)) {
                        yield return formula.Text!;
                    }
                }
            }
        }

        private static bool ContainsDefinedNameToken(string formula, string definedName) {
            for (int index = 0; index <= formula.Length - definedName.Length; index++) {
                if (!string.Equals(formula.Substring(index, definedName.Length), definedName, StringComparison.OrdinalIgnoreCase)) {
                    continue;
                }

                int afterIndex = index + definedName.Length;
                if ((index == 0 || !IsDefinedNameCharacter(formula[index - 1]))
                    && (afterIndex == formula.Length || !IsDefinedNameCharacter(formula[afterIndex]))) {
                    return true;
                }
            }

            return false;
        }

        private static bool IsDefinedNameCharacter(char value) {
            return char.IsLetterOrDigit(value) || value == '_' || value == '.' || value == '\\';
        }
    }
}
