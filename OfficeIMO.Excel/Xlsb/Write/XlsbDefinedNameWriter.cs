using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and writes common workbook defined names and their self references.</summary>
    internal static class XlsbDefinedNameWriter {
        private const int BrtName = 39;
        private const int BrtBeginExternals = 353;
        private const int BrtEndExternals = 354;
        private const int BrtSupSelf = 357;
        private const int BrtExternSheet = 362;

        internal static void Validate(DefinedNames? definedNames, IReadOnlyList<ExcelSheet> sheets) {
            CreatePlan(definedNames, sheets);
        }

        internal static void Write(Stream output, DefinedNames? definedNames, IReadOnlyList<ExcelSheet> sheets) {
            if (output == null) throw new ArgumentNullException(nameof(output));
            XlsbDefinedNamePlan plan = CreatePlan(definedNames, sheets);
            if (plan.Records.Count == 0) return;

            XlsbRecordWriter.Write(output, BrtBeginExternals);
            XlsbRecordWriter.Write(output, BrtSupSelf);
            XlsbRecordWriter.Write(output, BrtExternSheet, CreateExternalSheetPayload(sheets.Count));
            XlsbRecordWriter.Write(output, BrtEndExternals);
            foreach (XlsbDefinedNameRecord record in plan.Records) {
                XlsbRecordWriter.Write(output, BrtName, CreateNamePayload(record));
            }
        }

        private static XlsbDefinedNamePlan CreatePlan(DefinedNames? definedNames, IReadOnlyList<ExcelSheet> sheets) {
            if (sheets == null) throw new ArgumentNullException(nameof(sheets));
            if (definedNames != null
                && definedNames.HasChildren
                && definedNames.ChildElements.Any(element => element is not DefinedName)) {
                throw new NotSupportedException("Native XLSB generation supports only definedName children in the definedNames collection.");
            }
            if (definedNames != null) EnsureOnlyAttributes(definedNames);

            var records = new List<XlsbDefinedNameRecord>();
            var filterDatabaseSheets = new HashSet<uint>();
            foreach (DefinedName definedName in definedNames?.Elements<DefinedName>() ?? Enumerable.Empty<DefinedName>()) {
                EnsureOnlyAttributes(definedName, "name", "localSheetId", "hidden", "comment");
                if (definedName.HasChildren) {
                    throw new NotSupportedException("Native XLSB generation does not support child content in defined names.");
                }

                string openXmlName = definedName.Name?.Value ?? string.Empty;
                if (string.IsNullOrWhiteSpace(openXmlName) || openXmlName.Length > 255) {
                    throw new NotSupportedException("Native XLSB generation requires defined names between 1 and 255 characters.");
                }
                uint localSheetIndex = definedName.LocalSheetId?.Value ?? uint.MaxValue;
                if (localSheetIndex != uint.MaxValue && localSheetIndex >= sheets.Count) {
                    throw new NotSupportedException($"Defined name '{openXmlName}' is scoped outside the workbook.");
                }

                string? comment = definedName.Comment?.Value;
                if (comment != null && comment.Length >= 256) {
                    throw new NotSupportedException($"Defined name '{openXmlName}' has a comment longer than the XLSB limit of 255 characters.");
                }
                string formulaText = definedName.Text ?? string.Empty;
                if (!TryEncodeReferenceFormula(
                    formulaText,
                    localSheetIndex,
                    sheets,
                    out byte[]? formulaTokens,
                    out string? reason)) {
                    throw new NotSupportedException($"Native XLSB generation cannot write defined name '{openXmlName}': {reason}.");
                }

                bool builtIn = TryMapBuiltInName(openXmlName, out string recordName);
                if (!builtIn && openXmlName.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)) {
                    throw new NotSupportedException($"Native XLSB generation does not yet support built-in defined name '{openXmlName}'.");
                }
                uint flags = (definedName.Hidden?.Value == true ? 0x00000001U : 0U)
                    | (builtIn ? 0x00000020U : 0U);
                records.Add(new XlsbDefinedNameRecord(flags, localSheetIndex, recordName, formulaTokens!, comment));
                if (builtIn
                    && string.Equals(recordName, "_FilterDatabase", StringComparison.OrdinalIgnoreCase)
                    && localSheetIndex != uint.MaxValue) {
                    filterDatabaseSheets.Add(localSheetIndex);
                }
            }

            AppendAutoFilterDefinedNames(records, filterDatabaseSheets, sheets);

            return new XlsbDefinedNamePlan(records);
        }

        private static void AppendAutoFilterDefinedNames(
            List<XlsbDefinedNameRecord> records,
            HashSet<uint> existingSheets,
            IReadOnlyList<ExcelSheet> sheets) {
            for (int sheetIndex = 0; sheetIndex < sheets.Count; sheetIndex++) {
                uint localSheetIndex = checked((uint)sheetIndex);
                if (existingSheets.Contains(localSheetIndex)) continue;
                AutoFilter? autoFilter = sheets[sheetIndex].WorksheetPart.Worksheet?.GetFirstChild<AutoFilter>();
                if (autoFilter == null) continue;
                if (!XlsbWorksheetAutoFilterWriter.TryGetRange(autoFilter, out XlsbCellRange? range) || range == null) {
                    throw new NotSupportedException($"Native XLSB generation cannot encode the AutoFilter range on worksheet '{sheets[sheetIndex].Name}'.");
                }

                string escapedSheetName = sheets[sheetIndex].Name.Replace("'", "''");
                string formula = "'" + escapedSheetName + "'!" + range.ToA1Reference();
                if (!TryEncodeReferenceFormula(
                    formula,
                    localSheetIndex,
                    sheets,
                    out byte[]? formulaTokens,
                    out string? reason)) {
                    throw new NotSupportedException($"Native XLSB generation cannot create the AutoFilter defined name on worksheet '{sheets[sheetIndex].Name}': {reason}.");
                }
                records.Add(new XlsbDefinedNameRecord(
                    0x00000021U,
                    localSheetIndex,
                    "_FilterDatabase",
                    formulaTokens!,
                    comment: null));
            }
        }

        private static bool TryEncodeReferenceFormula(
            string formulaText,
            uint localSheetIndex,
            IReadOnlyList<ExcelSheet> sheets,
            out byte[]? formulaTokens,
            out string? reason) {
            formulaTokens = null;
            reason = null;
            string normalized = formulaText.Trim();
            if (normalized.StartsWith("=", StringComparison.Ordinal)) normalized = normalized.Substring(1).Trim();
            if (normalized.Length == 0) {
                reason = "its formula is empty";
                return false;
            }
            if (!TrySplitReferences(normalized, out IReadOnlyList<string>? references)) {
                reason = "its reference union is malformed";
                return false;
            }

            using var output = new MemoryStream();
            for (int index = 0; index < references!.Count; index++) {
                if (!TryParseReference(references[index], localSheetIndex, sheets, out int sheetIndex, out ReferenceBounds bounds, out reason)) {
                    return false;
                }
                if (sheetIndex > ushort.MaxValue - 2) {
                    reason = "it refers to a sheet beyond the XLSB external-sheet index limit";
                    return false;
                }

                ushort externalSheetIndex = checked((ushort)(sheetIndex + 2));
                WriteReferenceToken(output, externalSheetIndex, bounds);
                if (index > 0) output.WriteByte(0x10); // PtgUnion
            }
            if (output.Length == 0 || output.Length > 16_384) {
                reason = "its encoded formula exceeds XLSB limits";
                return false;
            }
            formulaTokens = output.ToArray();
            return true;
        }

        private static bool TrySplitReferences(string text, out IReadOnlyList<string>? references) {
            var result = new List<string>();
            int start = 0;
            bool quoted = false;
            for (int index = 0; index < text.Length; index++) {
                if (text[index] == '\'') {
                    if (quoted && index + 1 < text.Length && text[index + 1] == '\'') {
                        index++;
                        continue;
                    }
                    quoted = !quoted;
                } else if (text[index] == ',' && !quoted) {
                    string part = text.Substring(start, index - start).Trim();
                    if (part.Length == 0) {
                        references = null;
                        return false;
                    }
                    result.Add(part);
                    start = index + 1;
                }
            }
            if (quoted) {
                references = null;
                return false;
            }
            string finalPart = text.Substring(start).Trim();
            if (finalPart.Length == 0) {
                references = null;
                return false;
            }
            result.Add(finalPart);
            references = result;
            return true;
        }

        private static bool TryParseReference(
            string reference,
            uint localSheetIndex,
            IReadOnlyList<ExcelSheet> sheets,
            out int sheetIndex,
            out ReferenceBounds bounds,
            out string? reason) {
            sheetIndex = -1;
            bounds = default;
            reason = null;
            string localReference;
            if (SheetNameLookup.TryParseSheetQualifiedReference(
                reference,
                out string sheetName,
                out string parsedReference,
                allowExternalWorkbookReferences: false)) {
                for (int index = 0; index < sheets.Count; index++) {
                    if (SheetNameLookup.Matches(sheets[index].Name, sheetName)) {
                        sheetIndex = index;
                        break;
                    }
                }
                if (sheetIndex < 0) {
                    reason = "it refers to a sheet outside the workbook";
                    return false;
                }
                localReference = parsedReference;
            } else if (localSheetIndex != uint.MaxValue) {
                sheetIndex = checked((int)localSheetIndex);
                localReference = reference;
            } else {
                reason = "only internal sheet-qualified cell, range, print-area, and print-title references are supported";
                return false;
            }

            return TryParseLocalReference(localReference, out bounds, out reason);
        }

        private static bool TryParseLocalReference(string reference, out ReferenceBounds bounds, out string? reason) {
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

            string[] parts = normalized.Split(':');
            if (parts.Length == 2
                && int.TryParse(parts[0], out firstRow)
                && int.TryParse(parts[1], out lastRow)) {
                return TryCreateBounds(firstRow, 1, lastRow, 16_384, out bounds, out reason);
            }
            if (parts.Length == 2) {
                firstColumn = A1.ColumnLettersToIndex(parts[0]);
                lastColumn = A1.ColumnLettersToIndex(parts[1]);
                if (firstColumn > 0 && lastColumn > 0) {
                    return TryCreateBounds(1, firstColumn, 1_048_576, lastColumn, out bounds, out reason);
                }
            }

            reason = "its formula is outside the supported internal-reference subset";
            return false;
        }

        private static bool TryCreateBounds(
            int firstRow,
            int firstColumn,
            int lastRow,
            int lastColumn,
            out ReferenceBounds bounds,
            out string? reason) {
            bounds = default;
            reason = null;
            if (firstRow > lastRow) (firstRow, lastRow) = (lastRow, firstRow);
            if (firstColumn > lastColumn) (firstColumn, lastColumn) = (lastColumn, firstColumn);
            if (firstRow < 1 || lastRow > 1_048_576 || firstColumn < 1 || lastColumn > 16_384) {
                reason = "its range is outside XLSB worksheet limits";
                return false;
            }
            bounds = new ReferenceBounds(firstRow - 1, firstColumn - 1, lastRow - 1, lastColumn - 1);
            return true;
        }

        private static void WriteReferenceToken(Stream output, ushort externalSheetIndex, ReferenceBounds bounds) {
            if (bounds.IsSingleCell) {
                output.WriteByte(0x3A); // PtgRef3dV
                WriteUInt16(output, externalSheetIndex);
                WriteUInt32(output, checked((uint)bounds.FirstRow));
                WriteUInt16(output, checked((ushort)bounds.FirstColumn));
                return;
            }

            output.WriteByte(0x3B); // PtgArea3dV
            WriteUInt16(output, externalSheetIndex);
            WriteUInt32(output, checked((uint)bounds.FirstRow));
            WriteUInt32(output, checked((uint)bounds.LastRow));
            WriteUInt16(output, checked((ushort)bounds.FirstColumn));
            WriteUInt16(output, checked((ushort)bounds.LastColumn));
        }

        private static byte[] CreateNamePayload(XlsbDefinedNameRecord record) {
            using var output = new MemoryStream(32 + record.Name.Length * 2 + record.FormulaTokens.Length);
            WriteUInt32(output, record.Flags);
            output.WriteByte(0); // chKey
            WriteUInt32(output, record.LocalSheetIndex);
            WriteWideString(output, record.Name);
            WriteUInt32(output, checked((uint)record.FormulaTokens.Length));
            output.Write(record.FormulaTokens, 0, record.FormulaTokens.Length);
            WriteUInt32(output, 0); // cbExtra
            WriteNullableWideString(output, record.Comment);
            return output.ToArray();
        }

        private static byte[] CreateExternalSheetPayload(int sheetCount) {
            using var output = new MemoryStream(28 + sheetCount * 12);
            WriteUInt32(output, checked((uint)sheetCount + 2U));
            WriteExternalSheetReference(output, -2); // workbook-level reference
            WriteExternalSheetReference(output, -1); // #REF! reference
            for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
                WriteExternalSheetReference(output, sheetIndex);
            }
            return output.ToArray();
        }

        private static void WriteExternalSheetReference(Stream output, int sheetIndex) {
            WriteUInt32(output, 0); // first and only supporting link is BrtSupSelf
            WriteUInt32(output, unchecked((uint)sheetIndex));
            WriteUInt32(output, unchecked((uint)sheetIndex));
        }

        private static bool TryMapBuiltInName(string openXmlName, out string recordName) {
            if (string.Equals(openXmlName, "_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase)) {
                recordName = "Print_Area";
                return true;
            }
            if (string.Equals(openXmlName, "_xlnm.Print_Titles", StringComparison.OrdinalIgnoreCase)) {
                recordName = "Print_Titles";
                return true;
            }
            if (string.Equals(openXmlName, "_FilterDatabase", StringComparison.OrdinalIgnoreCase)
                || string.Equals(openXmlName, "_xlnm._FilterDatabase", StringComparison.OrdinalIgnoreCase)) {
                recordName = "_FilterDatabase";
                return true;
            }
            recordName = openXmlName;
            return false;
        }

        private static void EnsureOnlyAttributes(OpenXmlElement element, params string[] allowedNames) {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            OpenXmlAttribute? unsupported = element.GetAttributes()
                .Cast<OpenXmlAttribute?>()
                .FirstOrDefault(attribute => attribute.HasValue
                    && !string.Equals(attribute.Value.NamespaceUri, "http://www.w3.org/2000/xmlns/", StringComparison.Ordinal)
                    && !allowed.Contains(attribute.Value.LocalName));
            if (unsupported.HasValue) {
                throw new NotSupportedException($"Native XLSB generation does not yet support defined-name attribute '{unsupported.Value.LocalName}'.");
            }
        }

        private static void WriteWideString(Stream output, string value) {
            WriteUInt32(output, checked((uint)value.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
        }

        private static void WriteNullableWideString(Stream output, string? value) {
            if (value == null) {
                WriteUInt32(output, uint.MaxValue);
                return;
            }
            WriteWideString(output, value);
        }

        private static void WriteUInt16(Stream output, ushort value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
        }

        private static void WriteUInt32(Stream output, uint value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
            output.WriteByte((byte)(value >> 16));
            output.WriteByte((byte)(value >> 24));
        }

        private readonly struct ReferenceBounds {
            internal ReferenceBounds(int firstRow, int firstColumn, int lastRow, int lastColumn) {
                FirstRow = firstRow;
                FirstColumn = firstColumn;
                LastRow = lastRow;
                LastColumn = lastColumn;
            }

            internal int FirstRow { get; }
            internal int FirstColumn { get; }
            internal int LastRow { get; }
            internal int LastColumn { get; }
            internal bool IsSingleCell => FirstRow == LastRow && FirstColumn == LastColumn;
        }

        private sealed class XlsbDefinedNamePlan {
            internal static XlsbDefinedNamePlan Empty { get; } = new XlsbDefinedNamePlan(
                Array.Empty<XlsbDefinedNameRecord>());

            internal XlsbDefinedNamePlan(IReadOnlyList<XlsbDefinedNameRecord> records) {
                Records = records;
            }

            internal IReadOnlyList<XlsbDefinedNameRecord> Records { get; }
        }

        private sealed class XlsbDefinedNameRecord {
            internal XlsbDefinedNameRecord(uint flags, uint localSheetIndex, string name, byte[] formulaTokens, string? comment) {
                Flags = flags;
                LocalSheetIndex = localSheetIndex;
                Name = name;
                FormulaTokens = formulaTokens;
                Comment = comment;
            }

            internal uint Flags { get; }
            internal uint LocalSheetIndex { get; }
            internal string Name { get; }
            internal byte[] FormulaTokens { get; }
            internal string? Comment { get; }
        }
    }
}
