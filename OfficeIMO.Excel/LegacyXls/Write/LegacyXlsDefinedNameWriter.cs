using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal static class LegacyXlsDefinedNameWriter {
        internal static bool HasWorkbookDefinedNames(ExcelDocument document) {
            return document.WorkbookRoot.DefinedNames?.Elements<DefinedName>().Any() == true;
        }

        internal static bool SupportsWorkbookDefinedNames(
            ExcelDocument document,
            IReadOnlyList<ExcelSheet> sheets,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            out string? reason) {
            reason = null;
            DefinedNames? definedNames = document.WorkbookRoot.DefinedNames;
            if (definedNames == null) {
                return true;
            }

            foreach (DefinedName definedName in definedNames.Elements<DefinedName>()) {
                if (!TryCreateDefinedNamePayload(definedName, sheets, formulaNameIndex, out _, out reason)) {
                    return false;
                }
            }

            return true;
        }

        internal static IReadOnlyList<byte[]> CreateDefinedNamePayloads(
            ExcelDocument document,
            IReadOnlyList<ExcelSheet> sheets,
            LegacyXlsFormulaNameIndex formulaNameIndex) {
            DefinedNames? definedNames = document.WorkbookRoot.DefinedNames;
            if (definedNames == null) {
                return Array.Empty<byte[]>();
            }

            var payloads = new List<byte[]>();
            foreach (DefinedName definedName in definedNames.Elements<DefinedName>()) {
                if (TryCreateDefinedNamePayload(definedName, sheets, formulaNameIndex, out byte[]? payload, out _) && payload != null) {
                    payloads.Add(payload);
                }
            }

            return payloads;
        }

        internal static LegacyXlsFormulaNameIndex CreateFormulaNameIndex(
            ExcelDocument document,
            IReadOnlyList<ExcelSheet> sheets,
            LegacyXlsExternSheetTable externSheetTable) {
            DefinedNames? definedNames = document.WorkbookRoot.DefinedNames;
            if (definedNames == null) {
                return new LegacyXlsFormulaNameIndex(
                    new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase),
                    new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase),
                    externSheetTable);
            }

            var globalNames = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
            var localNames = new Dictionary<string, uint>(StringComparer.OrdinalIgnoreCase);
            uint oneBasedNameIndex = checked((uint)LegacyXlsAutoFilterWriter.CreateDefinedNamePayloads(sheets).Count + 1U);

            foreach (DefinedName definedName in definedNames.Elements<DefinedName>()) {
                if (!TryReserveDefinedNameIndexSlot(definedName, sheets, out ushort localSheetIndex)) {
                    continue;
                }

                string name = definedName.Name?.Value ?? string.Empty;
                if (!IsBuiltInName(name) && !name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)) {
                    if (localSheetIndex > 0) {
                        localNames[LegacyXlsFormulaNameIndex.CreateLocalKey(localSheetIndex - 1, name)] = oneBasedNameIndex;
                    } else {
                        globalNames[name] = oneBasedNameIndex;
                    }
                }

                oneBasedNameIndex++;
            }

            return new LegacyXlsFormulaNameIndex(globalNames, localNames, externSheetTable);
        }

        private static bool TryReserveDefinedNameIndexSlot(DefinedName definedName, IReadOnlyList<ExcelSheet> sheets, out ushort localSheetIndex) {
            localSheetIndex = 0;
            string name = definedName.Name?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(name) || name.Length > byte.MaxValue) {
                return false;
            }

            if (!TryGetLocalSheetIndex(definedName, sheets.Count, out localSheetIndex, out _)) {
                return false;
            }

            if (string.Equals(name, "_FilterDatabase", StringComparison.OrdinalIgnoreCase)
                && localSheetIndex > 0
                && HasWorksheetAutoFilter(sheets[localSheetIndex - 1])) {
                return false;
            }

            return true;
        }

        private static bool TryCreateDefinedNamePayload(
            DefinedName definedName,
            IReadOnlyList<ExcelSheet> sheets,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            out byte[]? payload,
            out string? reason) {
            payload = null;
            reason = null;
            string name = definedName.Name?.Value ?? string.Empty;
            if (string.IsNullOrWhiteSpace(name)) {
                reason = "defined names without names";
                return false;
            }

            if (name.Length > byte.MaxValue) {
                reason = "defined names longer than 255 characters";
                return false;
            }

            if (!TryGetLocalSheetIndex(definedName, sheets.Count, out ushort localSheetIndex, out reason)) {
                return false;
            }

            if (string.Equals(name, "_FilterDatabase", StringComparison.OrdinalIgnoreCase)
                && localSheetIndex > 0
                && HasWorksheetAutoFilter(sheets[localSheetIndex - 1])) {
                return true;
            }

            string reference = definedName.Text ?? string.Empty;
            if (!TryEncodeDefinedNameFormula(name, reference, sheets, formulaNameIndex, localSheetIndex, out byte[]? formula, out reason)) {
                return false;
            }

            ushort flags = definedName.Hidden?.Value == true ? (ushort)0x0001 : (ushort)0;
            string recordName = name;
            if (TryGetBuiltInNameCode(name, out char builtInCode)) {
                flags |= 0x0020;
                recordName = builtInCode.ToString();
            } else if (name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)) {
                reason = "unsupported built-in defined names";
                return false;
            }

            if (!SupportsDefinedNamePayload(recordName, formula!, out reason)) {
                return false;
            }

            payload = BuildDefinedNamePayload(recordName, formula!, localSheetIndex, flags);
            return true;
        }

        private static bool TryGetLocalSheetIndex(DefinedName definedName, int sheetCount, out ushort localSheetIndex, out string? reason) {
            localSheetIndex = 0;
            reason = null;
            if (definedName.LocalSheetId == null) {
                return true;
            }

            uint value = definedName.LocalSheetId.Value;
            if (value >= sheetCount || value > ushort.MaxValue - 1U) {
                reason = "defined names scoped outside the workbook";
                return false;
            }

            localSheetIndex = checked((ushort)(value + 1U));
            return true;
        }

        private static bool TryEncodeDefinedNameFormula(
            string name,
            string reference,
            IReadOnlyList<ExcelSheet> sheets,
            LegacyXlsFormulaNameIndex formulaNameIndex,
            ushort localSheetIndex,
            out byte[]? formula,
            out string? reason) {
            formula = null;
            reason = null;
            if (string.IsNullOrWhiteSpace(reference)) {
                reason = "defined names without references";
                return false;
            }

            if (string.Equals(name, "_xlnm.Print_Titles", StringComparison.OrdinalIgnoreCase)) {
                return TryEncodePrintTitlesFormula(reference, sheets, out formula, out reason);
            }

            int formulaSheetIndex = localSheetIndex > 0 ? localSheetIndex - 1 : -1;
            if (LegacyXlsFormulaEncoder.TryEncodeListSource(reference, formulaNameIndex, formulaSheetIndex, out byte[] encodedFormula, out reason)) {
                formula = encodedFormula;
                return true;
            }

            if (IsExternalWorkbookFormulaReason(reason)) {
                reason = "external workbook references in defined-name formulas";
            } else if (ContainsExternalDefinedNameReference(reference)) {
                reason = "external defined-name references";
            } else {
                reason = "defined-name formulas outside the supported native XLS formula subset";
            }

            return false;
        }

        private static bool IsExternalWorkbookFormulaReason(string? reason) {
            return reason?.IndexOf("external workbook references", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool ContainsExternalDefinedNameReference(string reference) {
            bool inStringLiteral = false;
            for (int i = 0; i < reference.Length; i++) {
                char ch = reference[i];
                if (ch == '"') {
                    if (inStringLiteral && i + 1 < reference.Length && reference[i + 1] == '"') {
                        i++;
                        continue;
                    }

                    inStringLiteral = !inStringLiteral;
                    continue;
                }

                if (!inStringLiteral && ch == '[') {
                    int close = reference.IndexOf(']', i + 1);
                    if (close > i && IsExternalDefinedNameWorkbookToken(reference.Substring(i + 1, close - i - 1))) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool IsExternalDefinedNameWorkbookToken(string token) {
            string trimmed = token.Trim();
            if (int.TryParse(trimmed, out _)) {
                return true;
            }

            return trimmed.IndexOf(".xls", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool TryEncodeSingleReferenceFormula(string reference, IReadOnlyList<ExcelSheet> sheets, out byte[]? formula, out string? reason) {
            formula = null;
            reason = null;
            if (!TryParseSheetQualifiedReference(reference, sheets, out int sheetIndex, out string? localReference, out reason)) {
                return false;
            }

            if (!TryParseCellOrRange(localReference!, out ReferenceBounds bounds, out reason)) {
                return false;
            }

            formula = bounds.IsSingleCell
                ? BuildRef3dFormula(checked((ushort)sheetIndex), bounds)
                : BuildArea3dFormula(checked((ushort)sheetIndex), bounds);
            return true;
        }

        private static bool TryEncodePrintTitlesFormula(string reference, IReadOnlyList<ExcelSheet> sheets, out byte[]? formula, out string? reason) {
            formula = null;
            reason = null;
            string[] parts = reference.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0 || parts.Length > 2) {
                reason = "print-title defined names with more than two ranges";
                return false;
            }

            var formulas = new List<byte[]>(parts.Length);
            int? expectedSheetIndex = null;
            foreach (string part in parts) {
                if (!TryParseSheetQualifiedReference(part, sheets, out int sheetIndex, out string? localReference, out reason)) {
                    return false;
                }

                if (expectedSheetIndex.HasValue && expectedSheetIndex.Value != sheetIndex) {
                    reason = "print-title defined names spanning multiple sheets";
                    return false;
                }

                expectedSheetIndex = sheetIndex;
                if (!TryParsePrintTitleReference(localReference!, out ReferenceBounds bounds, out reason)) {
                    return false;
                }

                formulas.Add(BuildArea3dFormula(checked((ushort)sheetIndex), bounds));
            }

            using var stream = new MemoryStream();
            foreach (byte[] partFormula in formulas) {
                stream.Write(partFormula, 0, partFormula.Length);
            }

            if (formulas.Count == 2) {
                stream.WriteByte(0x10);
            }

            formula = stream.ToArray();
            return true;
        }

        private static bool TryParseSheetQualifiedReference(string reference, IReadOnlyList<ExcelSheet> sheets, out int sheetIndex, out string? localReference, out string? reason) {
            sheetIndex = -1;
            localReference = null;
            reason = null;
            if (!SheetNameLookup.TryParseSheetQualifiedReference(reference, out string sheetName, out string parsedReference, allowExternalWorkbookReferences: false)) {
                reason = "defined names without sheet-qualified internal references";
                return false;
            }

            for (int i = 0; i < sheets.Count; i++) {
                if (SheetNameLookup.Matches(sheets[i].Name, sheetName)) {
                    sheetIndex = i;
                    localReference = parsedReference;
                    return true;
                }
            }

            reason = "defined names referencing sheets outside the workbook";
            return false;
        }

        private static bool TryParseCellOrRange(string reference, out ReferenceBounds bounds, out string? reason) {
            bounds = default;
            reason = null;
            string normalized = reference.Replace("$", string.Empty).Trim();
            if (A1.TryParseRange(normalized, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)) {
                return TryCreateBounds(firstRow, firstColumn, lastRow, lastColumn, out bounds, out reason);
            }

            (int row, int column) = A1.ParseCellRef(normalized);
            if (row > 0 && column > 0) {
                return TryCreateBounds(row, column, row, column, out bounds, out reason);
            }

            reason = "defined-name formulas outside the supported cell/range reference subset";
            return false;
        }

        private static bool TryParsePrintTitleReference(string reference, out ReferenceBounds bounds, out string? reason) {
            bounds = default;
            reason = null;
            string normalized = reference.Replace("$", string.Empty).Trim();
            string[] parts = normalized.Split(':');
            if (parts.Length != 2 || string.IsNullOrWhiteSpace(parts[0]) || string.IsNullOrWhiteSpace(parts[1])) {
                return TryParseCellOrRange(reference, out bounds, out reason);
            }

            if (int.TryParse(parts[0], out int firstRow)
                && int.TryParse(parts[1], out int lastRow)) {
                return TryCreateBounds(firstRow, 1, lastRow, 256, out bounds, out reason);
            }

            int firstColumn = A1.ColumnLettersToIndex(parts[0]);
            int lastColumn = A1.ColumnLettersToIndex(parts[1]);
            if (firstColumn > 0 && lastColumn > 0) {
                return TryCreateBounds(1, firstColumn, 65536, lastColumn, out bounds, out reason);
            }

            return TryParseCellOrRange(reference, out bounds, out reason);
        }

        private static bool TryCreateBounds(int firstRow, int firstColumn, int lastRow, int lastColumn, out ReferenceBounds bounds, out string? reason) {
            bounds = default;
            reason = null;
            if (firstRow > lastRow) (firstRow, lastRow) = (lastRow, firstRow);
            if (firstColumn > lastColumn) (firstColumn, lastColumn) = (lastColumn, firstColumn);
            if (firstRow < 1 || firstColumn < 1 || lastRow > 65536 || lastColumn > 256) {
                reason = "defined names outside BIFF8 worksheet limits";
                return false;
            }

            bounds = new ReferenceBounds(
                checked((ushort)(firstRow - 1)),
                checked((ushort)(firstColumn - 1)),
                checked((ushort)(lastRow - 1)),
                checked((ushort)(lastColumn - 1)));
            return true;
        }

        private static bool TryGetBuiltInNameCode(string name, out char code) {
            if (string.Equals(name, "_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase)) {
                code = (char)0x06;
                return true;
            }

            if (string.Equals(name, "_xlnm.Print_Titles", StringComparison.OrdinalIgnoreCase)) {
                code = (char)0x07;
                return true;
            }

            if (string.Equals(name, "_FilterDatabase", StringComparison.OrdinalIgnoreCase)) {
                code = (char)0x0d;
                return true;
            }

            code = '\0';
            return false;
        }

        private static bool IsBuiltInName(string name) {
            return TryGetBuiltInNameCode(name, out _);
        }

        private static bool HasWorksheetAutoFilter(ExcelSheet sheet) {
            return sheet.WorksheetPart.Worksheet?.Elements<AutoFilter>().Any() == true;
        }

        private static bool SupportsDefinedNamePayload(string name, byte[] formula, out string? reason) {
            reason = null;
            byte[] nameBytes = EncodeUnicodeString(name, out _);
            if (formula.Length > ushort.MaxValue || 14L + nameBytes.Length + formula.Length > ushort.MaxValue) {
                reason = "defined-name formula payload lengths outside BIFF8 limits";
                return false;
            }

            return true;
        }

        private static byte[] BuildDefinedNamePayload(string name, byte[] formula, ushort localSheetIndex, ushort flags) {
            byte[] nameBytes = EncodeUnicodeString(name, out byte stringFlags);
            using var stream = new MemoryStream();
            WriteUInt16(stream, flags);
            stream.WriteByte(0);
            stream.WriteByte(checked((byte)name.Length));
            WriteUInt16(stream, checked((ushort)formula.Length));
            WriteUInt16(stream, 0);
            WriteUInt16(stream, localSheetIndex);
            stream.WriteByte(0);
            stream.WriteByte(0);
            stream.WriteByte(0);
            stream.WriteByte(0);
            stream.WriteByte(stringFlags);
            stream.Write(nameBytes, 0, nameBytes.Length);
            stream.Write(formula, 0, formula.Length);
            return stream.ToArray();
        }

        private static byte[] BuildRef3dFormula(ushort externSheetIndex, ReferenceBounds bounds) {
            using var stream = new MemoryStream();
            stream.WriteByte(0x3a);
            WriteUInt16(stream, externSheetIndex);
            WriteUInt16(stream, bounds.FirstRow);
            WriteUInt16(stream, bounds.FirstColumn);
            return stream.ToArray();
        }

        private static byte[] BuildArea3dFormula(ushort externSheetIndex, ReferenceBounds bounds) {
            using var stream = new MemoryStream();
            stream.WriteByte(0x3b);
            WriteUInt16(stream, externSheetIndex);
            WriteUInt16(stream, bounds.FirstRow);
            WriteUInt16(stream, bounds.LastRow);
            WriteUInt16(stream, bounds.FirstColumn);
            WriteUInt16(stream, bounds.LastColumn);
            return stream.ToArray();
        }

        private static byte[] EncodeUnicodeString(string text, out byte flags) {
            if (CanUseCompressedString(text)) {
                flags = 0;
                return Encoding.ASCII.GetBytes(text);
            }

            flags = 1;
            return Encoding.Unicode.GetBytes(text);
        }

        private static bool CanUseCompressedString(string text) {
            for (int i = 0; i < text.Length; i++) {
                if (text[i] > 0x7f) {
                    return false;
                }
            }

            return true;
        }

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)(value & 0xff));
            stream.WriteByte((byte)((value >> 8) & 0xff));
        }

        private readonly struct ReferenceBounds {
            internal ReferenceBounds(ushort firstRow, ushort firstColumn, ushort lastRow, ushort lastColumn) {
                FirstRow = firstRow;
                FirstColumn = firstColumn;
                LastRow = lastRow;
                LastColumn = lastColumn;
            }

            internal ushort FirstRow { get; }
            internal ushort FirstColumn { get; }
            internal ushort LastRow { get; }
            internal ushort LastColumn { get; }
            internal bool IsSingleCell => FirstRow == LastRow && FirstColumn == LastColumn;
        }
    }
}
