using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal sealed class LegacyXlsExternSheetTable {
        private const string ExternalLinkPathRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath";

        private readonly Dictionary<string, ushort> _rangeIndexes;
        private readonly Dictionary<string, ushort> _externalSheetIndexes;
        private readonly Dictionary<string, ExternalNameIndex> _externalNameIndexes;

        private LegacyXlsExternSheetTable(
            string[] sheetNames,
            Dictionary<string, ushort> rangeIndexes,
            Dictionary<string, ushort> externalSheetIndexes,
            Dictionary<string, ExternalNameIndex> externalNameIndexes,
            IReadOnlyList<SupportingLinkRecord> supportingLinkRecords,
            byte[] payload) {
            SheetNames = sheetNames ?? throw new ArgumentNullException(nameof(sheetNames));
            _rangeIndexes = rangeIndexes ?? throw new ArgumentNullException(nameof(rangeIndexes));
            _externalSheetIndexes = externalSheetIndexes ?? throw new ArgumentNullException(nameof(externalSheetIndexes));
            _externalNameIndexes = externalNameIndexes ?? throw new ArgumentNullException(nameof(externalNameIndexes));
            SupportingLinkRecords = supportingLinkRecords ?? throw new ArgumentNullException(nameof(supportingLinkRecords));
            Payload = payload ?? throw new ArgumentNullException(nameof(payload));
        }

        internal static LegacyXlsExternSheetTable Empty { get; } = new(
            Array.Empty<string>(),
            new Dictionary<string, ushort>(StringComparer.Ordinal),
            new Dictionary<string, ushort>(StringComparer.OrdinalIgnoreCase),
            new Dictionary<string, ExternalNameIndex>(StringComparer.OrdinalIgnoreCase),
            Array.Empty<SupportingLinkRecord>(),
            Array.Empty<byte>());

        internal string[] SheetNames { get; }

        internal IReadOnlyList<SupportingLinkRecord> SupportingLinkRecords { get; }

        internal byte[] Payload { get; }

        internal static LegacyXlsExternSheetTable Create(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            if (sheets == null) throw new ArgumentNullException(nameof(sheets));

            string[] sheetNames = sheets.Select(sheet => sheet.Name).ToArray();
            IReadOnlyList<ExternalWorkbookReference> externalReferences = CollectExternalWorkbookReferences(document, sheets);
            var supportingLinkRecords = new List<SupportingLinkRecord>();
            if (sheets.Count > 0) {
                supportingLinkRecords.Add(new SupportingLinkRecord(0x01ae, BuildSelfSupBookPayload(checked((ushort)sheets.Count))));
            }

            if (externalReferences.Count > 0) {
                for (int i = 0; i < externalReferences.Count; i++) {
                    externalReferences[i].SetSupBookIndex(checked((ushort)(i + 1)));
                    supportingLinkRecords.Add(new SupportingLinkRecord(0x01ae, BuildExternalWorkbookSupBookPayload(externalReferences[i])));
                    uint oneBasedExternalNameIndex = 1;
                    foreach (ExternalDefinedNameReference externalName in externalReferences[i].ExternalNames) {
                        externalReferences[i].SetExternalNameIndex(externalName, oneBasedExternalNameIndex++);
                        supportingLinkRecords.Add(new SupportingLinkRecord(0x0023, BuildExternalNamePayload(
                            externalName.Name,
                            externalReferences[i].GetOneBasedSheetIndex(externalName.SheetName))));
                    }
                }
            }

            var entries = new List<ExternSheetEntry>(sheets.Count + externalReferences.Sum(reference => reference.SheetNames.Count + (reference.ExternalNames.Count > 0 ? 1 : 0)));
            var rangeIndexes = new Dictionary<string, ushort>(StringComparer.Ordinal);
            var externalSheetIndexes = new Dictionary<string, ushort>(StringComparer.OrdinalIgnoreCase);
            var externalNameIndexes = new Dictionary<string, ExternalNameIndex>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < sheets.Count; i++) {
                entries.Add(new ExternSheetEntry(0, checked((ushort)i), checked((ushort)i)));
            }

            foreach ((int firstSheetIndex, int lastSheetIndex) in EnumerateFormulaSheetRanges(document, sheets)) {
                string key = CreateRangeKey(firstSheetIndex, lastSheetIndex);
                if (rangeIndexes.ContainsKey(key)) {
                    continue;
                }

                ushort externSheetIndex = checked((ushort)entries.Count);
                rangeIndexes[key] = externSheetIndex;
                entries.Add(new ExternSheetEntry(0, checked((ushort)firstSheetIndex), checked((ushort)lastSheetIndex)));
            }

            foreach (ExternalWorkbookReference externalReference in externalReferences) {
                for (int sheetIndex = 0; sheetIndex < externalReference.SheetNames.Count; sheetIndex++) {
                    ushort externSheetIndex = checked((ushort)entries.Count);
                    externalSheetIndexes[CreateExternalSheetKey(externalReference.Target, externalReference.SheetNames[sheetIndex])] = externSheetIndex;
                    entries.Add(new ExternSheetEntry(
                        externalReference.SupBookIndex,
                        checked((ushort)sheetIndex),
                        checked((ushort)sheetIndex)));
                }

                if (externalReference.ExternalNames.Count > 0) {
                    ushort externSheetIndex = checked((ushort)entries.Count);
                    entries.Add(new ExternSheetEntry(externalReference.SupBookIndex, 0, 0));
                    foreach (ExternalDefinedNameReference externalName in externalReference.ExternalNames) {
                        externalNameIndexes[CreateExternalNameKey(externalReference.Target, externalName.SheetName, externalName.Name)] = new ExternalNameIndex(
                            externSheetIndex,
                            externalReference.GetExternalNameIndex(externalName));
                    }
                }
            }

            return new LegacyXlsExternSheetTable(sheetNames, rangeIndexes, externalSheetIndexes, externalNameIndexes, supportingLinkRecords, BuildPayload(entries));
        }

        internal static bool SupportsDeclaredExternalWorkbookLinks(WorkbookPart workbookPart, out string? reason) {
            if (workbookPart == null) throw new ArgumentNullException(nameof(workbookPart));

            return TryCollectDeclaredExternalWorkbookReferences(
                workbookPart,
                new Dictionary<string, ExternalWorkbookReference>(StringComparer.OrdinalIgnoreCase),
                out reason);
        }

        internal bool TryGetSheetIndex(string sheetName, out ushort externSheetIndex) {
            externSheetIndex = 0;
            if (string.IsNullOrWhiteSpace(sheetName)) {
                return false;
            }

            if (TryParseExternalSheetName(sheetName, out string? externalTarget, out string? externalSheetName)) {
                return _externalSheetIndexes.TryGetValue(CreateExternalSheetKey(externalTarget!, externalSheetName!), out externSheetIndex);
            }

            for (int i = 0; i < SheetNames.Length; i++) {
                if (SheetNameLookup.Matches(SheetNames[i], sheetName)) {
                    externSheetIndex = checked((ushort)i);
                    return true;
                }
            }

            return false;
        }

        internal bool TryGetSheetRangeIndex(string firstSheetName, string lastSheetName, out ushort externSheetIndex) {
            externSheetIndex = 0;
            if (!TryFindSheetIndex(firstSheetName, out int firstSheetIndex)
                || !TryFindSheetIndex(lastSheetName, out int lastSheetIndex)
                || lastSheetIndex < firstSheetIndex) {
                return false;
            }

            if (firstSheetIndex == lastSheetIndex) {
                externSheetIndex = checked((ushort)firstSheetIndex);
                return true;
            }

            return _rangeIndexes.TryGetValue(CreateRangeKey(firstSheetIndex, lastSheetIndex), out externSheetIndex);
        }

        internal bool TryGetExternalNameIndex(string target, string? sheetName, string name, out ushort externSheetIndex, out uint oneBasedNameIndex) {
            externSheetIndex = 0;
            oneBasedNameIndex = 0;
            if (string.IsNullOrWhiteSpace(target) || string.IsNullOrWhiteSpace(name)) {
                return false;
            }

            if (!_externalNameIndexes.TryGetValue(CreateExternalNameKey(target, sheetName, name), out ExternalNameIndex index)) {
                return false;
            }

            externSheetIndex = index.ExternSheetIndex;
            oneBasedNameIndex = index.OneBasedNameIndex;
            return true;
        }

        private static IReadOnlyList<ExternalWorkbookReference> CollectExternalWorkbookReferences(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets) {
            var references = new Dictionary<string, ExternalWorkbookReference>(StringComparer.OrdinalIgnoreCase);
            if (!TryCollectDeclaredExternalWorkbookReferences(document.WorkbookPartRoot, references, out string? reason)) {
                throw new NotSupportedException($"Native XLS saving does not support {reason ?? "external workbook link metadata"}.");
            }

            foreach (string formulaText in EnumerateFormulaTexts(document, sheets)) {
                foreach (string sheetName in EnumerateExternalSheetNames(formulaText)) {
                    if (!TryParseExternalSheetName(sheetName, out string? target, out string? externalSheetName)) {
                        continue;
                    }

                    if (!references.TryGetValue(target!, out ExternalWorkbookReference? reference)) {
                        reference = new ExternalWorkbookReference(target!);
                        references[target!] = reference;
                    }

                    reference.AddSheet(externalSheetName!);
                }

                foreach ((string target, string? sheetName, string externalName) in EnumerateExternalDefinedNames(formulaText)) {
                    if (!references.TryGetValue(target, out ExternalWorkbookReference? reference)) {
                        reference = new ExternalWorkbookReference(target);
                        references[target] = reference;
                    }

                    if (!string.IsNullOrWhiteSpace(sheetName)) {
                        reference.AddSheet(sheetName!);
                    }

                    reference.AddExternalName(externalName, sheetName);
                }
            }

            foreach (ExternalWorkbookReference reference in references.Values) {
                if (reference.ExternalNames.Count > 0 && reference.SheetNames.Count == 0) {
                    reference.AddSheet("Sheet1");
                }
            }

            return references.Values.ToArray();
        }

        private static bool TryCollectDeclaredExternalWorkbookReferences(
            WorkbookPart workbookPart,
            Dictionary<string, ExternalWorkbookReference> references,
            out string? reason) {
            reason = null;
            Workbook? workbook = workbookPart.Workbook;
            if (workbook == null) {
                return true;
            }

            ExternalReferences? externalReferences = workbook.ExternalReferences;
            List<ExternalReference> referenceElements = externalReferences?.Elements<ExternalReference>().ToList() ?? new List<ExternalReference>();
            HashSet<OpenXmlPart> referencedParts = new HashSet<OpenXmlPart>();
            foreach (ExternalReference referenceElement in referenceElements) {
                string? relationshipId = referenceElement.Id?.Value;
                if (string.IsNullOrWhiteSpace(relationshipId)) {
                    reason = "external workbook links";
                    return false;
                }

                OpenXmlPart part;
                try {
                    part = workbookPart.GetPartById(relationshipId!);
                } catch (ArgumentOutOfRangeException) {
                    reason = "external workbook links";
                    return false;
                }

                if (part is not ExternalWorkbookPart externalWorkbookPart) {
                    reason = "external workbook links";
                    return false;
                }

                referencedParts.Add(externalWorkbookPart);
                if (!TryCollectDeclaredExternalWorkbookReference(externalWorkbookPart, references, out reason)) {
                    return false;
                }
            }

            foreach (ExternalWorkbookPart externalWorkbookPart in workbookPart.GetPartsOfType<ExternalWorkbookPart>()) {
                if (!referencedParts.Contains(externalWorkbookPart)) {
                    reason = "external workbook links";
                    return false;
                }
            }

            return true;
        }

        private static bool TryCollectDeclaredExternalWorkbookReference(
            ExternalWorkbookPart externalWorkbookPart,
            Dictionary<string, ExternalWorkbookReference> references,
            out string? reason) {
            reason = null;
            ExternalLink? externalLink = externalWorkbookPart.ExternalLink;
            if (externalLink == null) {
                reason = "external workbook links";
                return false;
            }

            ExternalBook? externalBook = externalLink.GetFirstChild<ExternalBook>();
            if (externalBook == null || externalLink.ChildElements.Any(child => child is not ExternalBook)) {
                reason = "external workbook link metadata";
                return false;
            }

            if (externalBook.ChildElements.Any(child => child is not DocumentFormat.OpenXml.Spreadsheet.SheetNames && child is not ExternalDefinedNames)) {
                reason = "external workbook link metadata";
                return false;
            }

            string? relationshipId = externalBook.Id?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) {
                reason = "external workbook links";
                return false;
            }

            ExternalRelationship? relationship = externalWorkbookPart.ExternalRelationships
                .FirstOrDefault(candidate => string.Equals(candidate.Id, relationshipId, StringComparison.Ordinal));
            if (relationship == null
                || !string.Equals(relationship.RelationshipType, ExternalLinkPathRelationshipType, StringComparison.Ordinal)
                || string.IsNullOrWhiteSpace(relationship.Uri.OriginalString)) {
                reason = "external workbook links";
                return false;
            }

            string target = relationship.Uri.OriginalString;
            if (!IsValidExternalWorkbookToken(target)) {
                reason = "external workbook links";
                return false;
            }

            if (!references.TryGetValue(target, out ExternalWorkbookReference? reference)) {
                reference = new ExternalWorkbookReference(target);
                references[target] = reference;
            }

            List<string> sheetNames = new List<string>();
            SheetNames? declaredSheetNames = externalBook.GetFirstChild<SheetNames>();
            if (declaredSheetNames != null) {
                foreach (SheetName sheetName in declaredSheetNames.Elements<SheetName>()) {
                    string? value = sheetName.Val?.Value;
                    if (!IsValidExternalSheetToken(value)) {
                        reason = "external workbook sheet names";
                        return false;
                    }

                    sheetNames.Add(value!);
                    reference.AddSheet(value!);
                }
            }

            ExternalDefinedNames? declaredNames = externalBook.GetFirstChild<ExternalDefinedNames>();
            if (declaredNames != null) {
                foreach (ExternalDefinedName definedName in declaredNames.Elements<ExternalDefinedName>()) {
                    string? name = definedName.Name?.Value;
                    if (!IsValidExternalNameOperand(name ?? string.Empty) || IsBuiltInExternalDefinedName(name!)) {
                        reason = "external workbook defined names";
                        return false;
                    }

                    string? localSheetName = null;
                    if (definedName.SheetId?.HasValue == true) {
                        uint sheetId = definedName.SheetId.Value;
                        if (sheetId >= sheetNames.Count || sheetId > ushort.MaxValue) {
                            reason = "external workbook defined names";
                            return false;
                        }

                        localSheetName = sheetNames[(int)sheetId];
                    }

                    if (!string.IsNullOrWhiteSpace(localSheetName)) {
                        reference.AddSheet(localSheetName!);
                    }

                    reference.AddExternalName(name!, localSheetName);
                }
            }

            return true;
        }

        private static IEnumerable<string> EnumerateFormulaTexts(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets) {
            DefinedNames? definedNames = document.WorkbookRoot.DefinedNames;
            if (definedNames != null) {
                foreach (DefinedName definedName in definedNames.Elements<DefinedName>()) {
                    yield return definedName.Text ?? string.Empty;
                }
            }

            foreach (ExcelSheet sheet in sheets) {
                Worksheet? worksheet = sheet.WorksheetPart.Worksheet;
                if (worksheet == null) {
                    continue;
                }

                DataValidations? validations = worksheet.GetFirstChild<DataValidations>();
                if (validations != null) {
                    foreach (DataValidation validation in validations.Elements<DataValidation>()) {
                        foreach (Formula1 formula in validation.Elements<Formula1>()) {
                            yield return formula.Text ?? string.Empty;
                        }

                        foreach (Formula2 formula in validation.Elements<Formula2>()) {
                            yield return formula.Text ?? string.Empty;
                        }
                    }
                }

                foreach (ConditionalFormatting conditionalFormatting in worksheet.Elements<ConditionalFormatting>()) {
                    foreach (ConditionalFormattingRule rule in conditionalFormatting.Elements<ConditionalFormattingRule>()) {
                        foreach (Formula formula in rule.Elements<Formula>()) {
                            yield return formula.Text ?? string.Empty;
                        }
                    }
                }

                SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData == null) {
                    continue;
                }

                foreach (Cell cell in sheetData.Descendants<Cell>()) {
                    yield return cell.CellFormula?.Text ?? string.Empty;
                }
            }
        }

        private static IEnumerable<(string Target, string? SheetName, string Name)> EnumerateExternalDefinedNames(string formulaText) {
            if (formulaText.IndexOf('[') < 0 || formulaText.IndexOf(']') < 0) {
                yield break;
            }

            foreach ((string target, string sheetName, string name) in EnumerateSheetScopedExternalDefinedNames(formulaText)) {
                yield return (target, sheetName, name);
            }

            bool inStringLiteral = false;
            for (int i = 0; i < formulaText.Length; i++) {
                char ch = formulaText[i];
                if (ch == '"') {
                    if (inStringLiteral && i + 1 < formulaText.Length && formulaText[i + 1] == '"') {
                        i++;
                        continue;
                    }

                    inStringLiteral = !inStringLiteral;
                    continue;
                }

                if (inStringLiteral || ch != '[') {
                    continue;
                }

                int close = formulaText.IndexOf(']', i + 1);
                if (close <= i + 1 || close >= formulaText.Length - 1) {
                    continue;
                }

                string target = formulaText.Substring(i + 1, close - i - 1).Trim();
                if (target.Length == 0 || target.Length > byte.MaxValue) {
                    continue;
                }

                int nameStart = close + 1;
                int nameEnd = nameStart;
                while (nameEnd < formulaText.Length && !IsExternalNameBoundary(formulaText[nameEnd])) {
                    nameEnd++;
                }

                if (nameEnd == nameStart) {
                    continue;
                }

                string name = formulaText.Substring(nameStart, nameEnd - nameStart).Trim();
                if (IsValidExternalNameOperand(name)) {
                    yield return (target, null, name);
                }
            }
        }

        private static IEnumerable<(string Target, string SheetName, string Name)> EnumerateSheetScopedExternalDefinedNames(string formulaText) {
            if (formulaText.IndexOf('!') < 0 || formulaText.IndexOf('[') < 0) {
                yield break;
            }

            for (int i = 0; i < formulaText.Length; i++) {
                if (formulaText[i] != '!') {
                    continue;
                }

                string sheetToken = ReadSheetTokenBeforeBang(formulaText, i);
                if (sheetToken.Length == 0) {
                    continue;
                }

                string sheetName = UnquoteSheetToken(sheetToken);
                if (!TryParseExternalSheetName(sheetName, out string? target, out string? externalSheetName)) {
                    continue;
                }

                int nameStart = i + 1;
                while (nameStart < formulaText.Length && char.IsWhiteSpace(formulaText[nameStart])) {
                    nameStart++;
                }

                int nameEnd = nameStart;
                while (nameEnd < formulaText.Length && !IsExternalNameBoundary(formulaText[nameEnd])) {
                    nameEnd++;
                }

                if (nameEnd == nameStart) {
                    continue;
                }

                string name = formulaText.Substring(nameStart, nameEnd - nameStart).Trim();
                if (IsValidExternalNameOperand(name)) {
                    yield return (target!, externalSheetName!, name);
                }
            }
        }

        private static IEnumerable<(int FirstSheetIndex, int LastSheetIndex)> EnumerateFormulaSheetRanges(ExcelDocument document, IReadOnlyList<ExcelSheet> sheets) {
            DefinedNames? definedNames = document.WorkbookRoot.DefinedNames;
            if (definedNames != null) {
                foreach (DefinedName definedName in definedNames.Elements<DefinedName>()) {
                    foreach ((int firstSheetIndex, int lastSheetIndex) in EnumerateFormulaSheetRanges(definedName.Text ?? string.Empty, sheets)) {
                        yield return (firstSheetIndex, lastSheetIndex);
                    }
                }
            }

            foreach (ExcelSheet sheet in sheets) {
                Worksheet? worksheet = sheet.WorksheetPart.Worksheet;
                if (worksheet == null) {
                    continue;
                }

                DataValidations? validations = worksheet.GetFirstChild<DataValidations>();
                if (validations != null) {
                    foreach (DataValidation validation in validations.Elements<DataValidation>()) {
                        foreach ((int firstSheetIndex, int lastSheetIndex) in EnumerateDataValidationFormulaSheetRanges(validation, sheets)) {
                            yield return (firstSheetIndex, lastSheetIndex);
                        }
                    }
                }

                foreach (ConditionalFormatting conditionalFormatting in worksheet.Elements<ConditionalFormatting>()) {
                    foreach ((int firstSheetIndex, int lastSheetIndex) in EnumerateConditionalFormattingFormulaSheetRanges(conditionalFormatting, sheets)) {
                        yield return (firstSheetIndex, lastSheetIndex);
                    }
                }

                SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData == null) {
                    continue;
                }

                foreach (Cell cell in sheetData.Descendants<Cell>()) {
                    string formulaText = cell.CellFormula?.Text ?? string.Empty;
                    foreach ((int firstSheetIndex, int lastSheetIndex) in EnumerateFormulaSheetRanges(formulaText, sheets)) {
                        yield return (firstSheetIndex, lastSheetIndex);
                    }
                }
            }
        }

        private static IEnumerable<(int FirstSheetIndex, int LastSheetIndex)> EnumerateConditionalFormattingFormulaSheetRanges(ConditionalFormatting conditionalFormatting, IReadOnlyList<ExcelSheet> sheets) {
            foreach (ConditionalFormattingRule rule in conditionalFormatting.Elements<ConditionalFormattingRule>()) {
                foreach (Formula formula in rule.Elements<Formula>()) {
                    foreach ((int firstSheetIndex, int lastSheetIndex) in EnumerateFormulaSheetRanges(formula.Text ?? string.Empty, sheets)) {
                        yield return (firstSheetIndex, lastSheetIndex);
                    }
                }
            }
        }

        private static IEnumerable<(int FirstSheetIndex, int LastSheetIndex)> EnumerateDataValidationFormulaSheetRanges(DataValidation validation, IReadOnlyList<ExcelSheet> sheets) {
            foreach (Formula1 formula in validation.Elements<Formula1>()) {
                foreach ((int firstSheetIndex, int lastSheetIndex) in EnumerateFormulaSheetRanges(formula.Text ?? string.Empty, sheets)) {
                    yield return (firstSheetIndex, lastSheetIndex);
                }
            }

            foreach (Formula2 formula in validation.Elements<Formula2>()) {
                foreach ((int firstSheetIndex, int lastSheetIndex) in EnumerateFormulaSheetRanges(formula.Text ?? string.Empty, sheets)) {
                    yield return (firstSheetIndex, lastSheetIndex);
                }
            }
        }

        private static IEnumerable<(int FirstSheetIndex, int LastSheetIndex)> EnumerateFormulaSheetRanges(string formulaText, IReadOnlyList<ExcelSheet> sheets) {
            if (formulaText.IndexOf('!') < 0) {
                yield break;
            }

            foreach (string sheetRangeName in EnumerateSheetRangeNames(formulaText)) {
                if (TryResolveSheetRange(sheetRangeName, sheets, out int firstSheetIndex, out int lastSheetIndex)) {
                    yield return (firstSheetIndex, lastSheetIndex);
                }
            }
        }

        private static IEnumerable<string> EnumerateExternalSheetNames(string formulaText) {
            if (formulaText.IndexOf('!') < 0 || formulaText.IndexOf('[') < 0) {
                yield break;
            }

            for (int i = 0; i < formulaText.Length; i++) {
                if (formulaText[i] != '!') {
                    continue;
                }

                string sheetToken = ReadSheetTokenBeforeBang(formulaText, i);
                if (sheetToken.Length == 0 || sheetToken.IndexOf('[') < 0 || sheetToken.IndexOf(']') < 0) {
                    continue;
                }

                string sheetName = UnquoteSheetToken(sheetToken);
                if (TryParseExternalSheetName(sheetName, out _, out _)) {
                    yield return sheetName;
                }
            }
        }

        internal static bool TryParseExternalSheetName(string sheetName, out string? target, out string? externalSheetName) {
            target = null;
            externalSheetName = null;
            string trimmed = sheetName.Trim();
            if (trimmed.Length < 4 || trimmed[0] != '[') {
                return false;
            }

            int close = trimmed.IndexOf(']');
            if (close <= 1 || close >= trimmed.Length - 1) {
                return false;
            }

            string parsedTarget = trimmed.Substring(1, close - 1).Trim();
            string parsedSheetName = trimmed.Substring(close + 1).Trim();
            if (parsedTarget.Length == 0
                || parsedSheetName.Length == 0
                || parsedTarget.Length > byte.MaxValue
                || parsedSheetName.Length > byte.MaxValue
                || parsedSheetName.IndexOf('[') >= 0
                || parsedSheetName.IndexOf(']') >= 0
                || parsedSheetName.IndexOf(':') >= 0) {
                return false;
            }

            target = parsedTarget;
            externalSheetName = parsedSheetName;
            return true;
        }

        private static IEnumerable<string> EnumerateSheetRangeNames(string formulaText) {
            for (int i = 0; i < formulaText.Length; i++) {
                if (formulaText[i] != '!') {
                    continue;
                }

                string sheetToken = ReadSheetTokenBeforeBang(formulaText, i);
                if (sheetToken.Length == 0 || sheetToken.IndexOf('[') >= 0 || sheetToken.IndexOf(']') >= 0) {
                    continue;
                }

                string sheetName = UnquoteSheetToken(sheetToken);
                if (sheetName.IndexOf(':') > 0) {
                    yield return sheetName;
                }
            }
        }

        private static string ReadSheetTokenBeforeBang(string formulaText, int bangIndex) {
            int end = bangIndex - 1;
            while (end >= 0 && char.IsWhiteSpace(formulaText[end])) {
                end--;
            }

            if (end < 0) {
                return string.Empty;
            }

            if (formulaText[end] == '\'') {
                int start = FindOpeningQuote(formulaText, end - 1);
                return start >= 0 ? formulaText.Substring(start, end - start + 1).Trim() : string.Empty;
            }

            int tokenStart = end;
            while (tokenStart >= 0 && !IsFormulaSheetTokenBoundary(formulaText[tokenStart])) {
                tokenStart--;
            }

            return formulaText.Substring(tokenStart + 1, end - tokenStart).Trim();
        }

        private static int FindOpeningQuote(string formulaText, int startIndex) {
            for (int i = startIndex; i >= 0; i--) {
                if (formulaText[i] != '\'') {
                    continue;
                }

                bool escapedQuote = i > 0 && formulaText[i - 1] == '\'';
                if (escapedQuote) {
                    i--;
                    continue;
                }

                return i;
            }

            return -1;
        }

        private static bool IsFormulaSheetTokenBoundary(char ch) {
            return char.IsWhiteSpace(ch)
                || ch == '('
                || ch == ')'
                || ch == ','
                || ch == '+'
                || ch == '-'
                || ch == '*'
                || ch == '/'
                || ch == '^'
                || ch == '&'
                || ch == '='
                || ch == '<'
                || ch == '>';
        }

        private static bool IsExternalNameBoundary(char ch) {
            return char.IsWhiteSpace(ch)
                || ch == '('
                || ch == ')'
                || ch == ','
                || ch == '+'
                || ch == '-'
                || ch == '*'
                || ch == '/'
                || ch == '^'
                || ch == '&'
                || ch == '='
                || ch == '<'
                || ch == '>'
                || ch == '{'
                || ch == '}';
        }

        private static bool IsValidExternalNameOperand(string name) {
            return name.Length > 0
                && name.Length <= byte.MaxValue
                && name.IndexOf('!') < 0
                && name.IndexOf(':') < 0
                && name.IndexOf('[') < 0
                && name.IndexOf(']') < 0
                && name.IndexOf('$') < 0;
        }

        private static bool IsValidExternalWorkbookToken(string? target) {
            return !string.IsNullOrWhiteSpace(target)
                && target!.Length <= byte.MaxValue
                && target.IndexOf('[') < 0
                && target.IndexOf(']') < 0;
        }

        private static bool IsValidExternalSheetToken(string? sheetName) {
            return !string.IsNullOrWhiteSpace(sheetName)
                && sheetName!.Length <= byte.MaxValue
                && sheetName.IndexOf('[') < 0
                && sheetName.IndexOf(']') < 0
                && sheetName.IndexOf(':') < 0;
        }

        private static bool IsBuiltInExternalDefinedName(string name) {
            return name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, "_FilterDatabase", StringComparison.OrdinalIgnoreCase);
        }

        private static string UnquoteSheetToken(string sheetToken) {
            string trimmed = sheetToken.Trim();
            if (trimmed.Length >= 2 && trimmed[0] == '\'' && trimmed[trimmed.Length - 1] == '\'') {
                return trimmed.Substring(1, trimmed.Length - 2).Replace("''", "'");
            }

            return trimmed;
        }

        private static bool TryResolveSheetRange(
            string sheetRangeName,
            IReadOnlyList<ExcelSheet> sheets,
            out int firstSheetIndex,
            out int lastSheetIndex) {
            firstSheetIndex = -1;
            lastSheetIndex = -1;
            string[] parts = sheetRangeName.Split(':');
            if (parts.Length != 2
                || string.IsNullOrWhiteSpace(parts[0])
                || string.IsNullOrWhiteSpace(parts[1])) {
                return false;
            }

            if (!TryFindSheetIndex(sheets, parts[0], out firstSheetIndex)
                || !TryFindSheetIndex(sheets, parts[1], out lastSheetIndex)
                || lastSheetIndex < firstSheetIndex) {
                return false;
            }

            return true;
        }

        private bool TryFindSheetIndex(string sheetName, out int sheetIndex) {
            sheetIndex = -1;
            for (int i = 0; i < SheetNames.Length; i++) {
                if (SheetNameLookup.Matches(SheetNames[i], sheetName)) {
                    sheetIndex = i;
                    return true;
                }
            }

            return false;
        }

        private static bool TryFindSheetIndex(IReadOnlyList<ExcelSheet> sheets, string sheetName, out int sheetIndex) {
            sheetIndex = -1;
            for (int i = 0; i < sheets.Count; i++) {
                if (SheetNameLookup.Matches(sheets[i].Name, sheetName)) {
                    sheetIndex = i;
                    return true;
                }
            }

            return false;
        }

        private static string CreateRangeKey(int firstSheetIndex, int lastSheetIndex) {
            return firstSheetIndex.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + "\0"
                + lastSheetIndex.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        private static string CreateExternalSheetKey(string target, string sheetName) {
            return target + "\0" + sheetName;
        }

        private static string CreateExternalNameKey(string target, string? sheetName, string name) {
            return target + "\0" + (sheetName ?? string.Empty) + "\0" + name;
        }

        private static byte[] BuildSelfSupBookPayload(ushort sheetCount) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, sheetCount);
            WriteUInt16(stream, 0x0401);
            return stream.ToArray();
        }

        private static byte[] BuildExternalWorkbookSupBookPayload(ExternalWorkbookReference reference) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)reference.SheetNames.Count));
            WriteUInt16(stream, checked((ushort)reference.Target.Length));
            WriteUnicodeStringNoCch(stream, reference.Target);
            foreach (string sheetName in reference.SheetNames) {
                WriteUnicodeString(stream, sheetName);
            }

            return stream.ToArray();
        }

        private static byte[] BuildExternalNamePayload(string name, ushort oneBasedSheetIndex) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, 0);
            WriteUInt16(stream, oneBasedSheetIndex);
            WriteUInt16(stream, 0);
            stream.WriteByte(checked((byte)name.Length));
            WriteUnicodeStringNoCch(stream, name);
            WriteUInt16(stream, 0);
            return stream.ToArray();
        }

        private static byte[] BuildPayload(IReadOnlyList<ExternSheetEntry> entries) {
            using var stream = new MemoryStream();
            WriteUInt16(stream, checked((ushort)entries.Count));
            foreach (ExternSheetEntry entry in entries) {
                WriteUInt16(stream, entry.SupBookIndex);
                WriteUInt16(stream, entry.FirstSheetIndex);
                WriteUInt16(stream, entry.LastSheetIndex);
            }

            return stream.ToArray();
        }

        private static void WriteUnicodeString(Stream stream, string value) {
            WriteUInt16(stream, checked((ushort)value.Length));
            WriteUnicodeStringNoCch(stream, value);
        }

        private static void WriteUnicodeStringNoCch(Stream stream, string value) {
            stream.WriteByte(1);
            byte[] encoded = System.Text.Encoding.Unicode.GetBytes(value);
            stream.Write(encoded, 0, encoded.Length);
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private readonly struct ExternSheetEntry {
            internal ExternSheetEntry(ushort supBookIndex, ushort firstSheetIndex, ushort lastSheetIndex) {
                SupBookIndex = supBookIndex;
                FirstSheetIndex = firstSheetIndex;
                LastSheetIndex = lastSheetIndex;
            }

            internal ushort SupBookIndex { get; }

            internal ushort FirstSheetIndex { get; }

            internal ushort LastSheetIndex { get; }
        }

        internal readonly struct SupportingLinkRecord {
            internal SupportingLinkRecord(ushort recordType, byte[] payload) {
                RecordType = recordType;
                Payload = payload ?? throw new ArgumentNullException(nameof(payload));
            }

            internal ushort RecordType { get; }

            internal byte[] Payload { get; }
        }

        private readonly struct ExternalNameIndex {
            internal ExternalNameIndex(ushort externSheetIndex, uint oneBasedNameIndex) {
                ExternSheetIndex = externSheetIndex;
                OneBasedNameIndex = oneBasedNameIndex;
            }

            internal ushort ExternSheetIndex { get; }

            internal uint OneBasedNameIndex { get; }
        }

        private sealed class ExternalWorkbookReference {
            private readonly List<string> _sheetNames = new List<string>();
            private readonly HashSet<string> _seenSheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            private readonly List<ExternalDefinedNameReference> _externalNames = new List<ExternalDefinedNameReference>();
            private readonly HashSet<string> _seenExternalNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            private readonly Dictionary<string, uint> _externalNameIndexes = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);

            internal ExternalWorkbookReference(string target) {
                Target = target;
            }

            internal string Target { get; }

            internal IReadOnlyList<string> SheetNames => _sheetNames;

            internal IReadOnlyList<ExternalDefinedNameReference> ExternalNames => _externalNames;

            internal ushort SupBookIndex { get; private set; }

            internal void AddSheet(string sheetName) {
                if (_seenSheetNames.Add(sheetName)) {
                    _sheetNames.Add(sheetName);
                }
            }

            internal void AddExternalName(string externalName, string? sheetName) {
                string key = CreateExternalNameKey(Target, sheetName, externalName);
                if (_seenExternalNames.Add(key)) {
                    _externalNames.Add(new ExternalDefinedNameReference(externalName, sheetName));
                }
            }

            internal void SetExternalNameIndex(ExternalDefinedNameReference externalName, uint oneBasedNameIndex) {
                _externalNameIndexes[CreateExternalNameKey(Target, externalName.SheetName, externalName.Name)] = oneBasedNameIndex;
            }

            internal uint GetExternalNameIndex(ExternalDefinedNameReference externalName) {
                return _externalNameIndexes[CreateExternalNameKey(Target, externalName.SheetName, externalName.Name)];
            }

            internal ushort GetOneBasedSheetIndex(string? sheetName) {
                if (string.IsNullOrWhiteSpace(sheetName)) {
                    return 0;
                }

                for (int i = 0; i < _sheetNames.Count; i++) {
                    if (SheetNameLookup.Matches(_sheetNames[i], sheetName!)) {
                        return checked((ushort)(i + 1));
                    }
                }

                return 0;
            }

            internal void SetSupBookIndex(ushort supBookIndex) {
                SupBookIndex = supBookIndex;
            }
        }

        private readonly struct ExternalDefinedNameReference {
            internal ExternalDefinedNameReference(string name, string? sheetName) {
                Name = name;
                SheetName = string.IsNullOrWhiteSpace(sheetName) ? null : sheetName;
            }

            internal string Name { get; }

            internal string? SheetName { get; }
        }
    }
}
