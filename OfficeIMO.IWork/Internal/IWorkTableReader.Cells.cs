using System.Globalization;
using System.Numerics;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    private static IWorkTableCell DecodeCell(byte[] buffer, int offset, int endOffset,
        int row, int column,
        IReadOnlyDictionary<uint, string> strings, IWorkTableRichTextCatalog richStrings,
        IReadOnlyDictionary<uint, IWorkWireMessage> formulas,
        IWorkReadOptions options, IWorkProjectionBudget projectionBudget,
        HashSet<uint> formulaRichStringIdentifiers, HashSet<uint> nonFormulaRichStringIdentifiers) {
        if (offset < 0 || endOffset < offset || endOffset > buffer.Length
            || offset > endOffset - 12) return Error(row, column, "Truncated cell record.");
        int version = buffer[offset];
        int type = buffer[offset + 1];
        if (version != 5) return Error(row, column, $"Unsupported cell storage version {version}.");
        uint flags = IWorkProtobuf.ReadUInt32(buffer, offset + 8);
        IWorkTableCell cell = DecodeModernCell(buffer, offset, endOffset, row, column, type, flags,
            strings, richStrings, formulas, options, projectionBudget,
            formulaRichStringIdentifiers, nonFormulaRichStringIdentifiers);
        // Only a complete, supported header establishes formula presence. Keep decode errors
        // as errors instead of inventing an expression or a recovered cache.
        return cell.Kind == IWorkCellKind.Error && (flags & (1u << 9)) != 0
            ? new IWorkTableCell(row, column, IWorkCellKind.Error, null,
                error: cell.Error, sourceFormulaIsDeclared: true,
                hasDecodeError: cell.HasDecodeError)
            : cell;
    }

    private static IWorkTableCell DecodeModernCell(byte[] buffer, int offset, int endOffset,
        int row, int column, int type, uint flags,
        IReadOnlyDictionary<uint, string> strings, IWorkTableRichTextCatalog richStrings,
        IReadOnlyDictionary<uint, IWorkWireMessage> formulas,
        IWorkReadOptions options, IWorkProjectionBudget projectionBudget,
        HashSet<uint> formulaRichStringIdentifiers, HashSet<uint> nonFormulaRichStringIdentifiers) {
        if ((flags & ~RecognizedCellValueMask) != 0) {
            return Error(row, column, "Cell storage contains unsupported value fields.");
        }
        int position = offset + 12;
        double? decimalValue = null;
        double doubleValue = 0;
        double dateValue = 0;
        uint stringIdentifier = 0;
        uint richStringIdentifier = 0;
        uint formulaIdentifier = 0;
        bool hasDecimal = false;
        bool hasDouble = false;
        bool hasDate = false;
        bool hasString = false;
        bool hasRichString = false;
        bool hasFormula = false;
        for (int bit = 0; bit < 21; bit++) {
            if ((flags & (1u << bit)) == 0) continue;
            int size = bit == 0 ? 16 : bit is 1 or 2 ? 8 : 4;
            if (position < 0 || position > endOffset - size) return Error(row, column, "Truncated cell value field.");
            switch (bit) {
                case 0:
                    decimalValue = ReadDecimal128(buffer, position);
                    hasDecimal = true;
                    break;
                case 1:
                    doubleValue = ReadDouble(buffer, position);
                    hasDouble = true;
                    break;
                case 2:
                    dateValue = ReadDouble(buffer, position);
                    hasDate = true;
                    break;
                case 3:
                    stringIdentifier = IWorkProtobuf.ReadUInt32(buffer, position);
                    hasString = true;
                    break;
                case 4:
                    richStringIdentifier = IWorkProtobuf.ReadUInt32(buffer, position);
                    hasRichString = true;
                    break;
                case 9:
                    formulaIdentifier = IWorkProtobuf.ReadUInt32(buffer, position);
                    hasFormula = true;
                    break;
            }
            position += size;
        }

        bool hasConflictingValueFields = type switch {
            0 => hasDecimal || hasDouble || hasDate || hasString || hasRichString || hasFormula,
            2 or 10 => hasDecimal && hasDouble || hasDate || hasString || hasRichString,
            3 => hasDecimal || hasDouble || hasDate || hasRichString,
            5 => hasDecimal || hasDouble || hasString || hasRichString,
            6 or 7 => hasDecimal || hasDate || hasString || hasRichString,
            8 => hasDecimal || hasDouble || hasDate || hasString || hasRichString,
            9 => hasDecimal || hasDouble || hasDate || hasString,
            _ => false
        };
        if (hasConflictingValueFields) {
            return Error(row, column, "Cell storage declares conflicting value fields.");
        }

        switch (type) {
            case 0:
                return new IWorkTableCell(row, column, IWorkCellKind.Empty, null);
            case 2:
            case 10:
                if (hasDecimal) {
                    return decimalValue.HasValue
                        ? FiniteNumber(row, column, decimalValue.Value, hasFormula,
                            formulaIdentifier, formulas, options, projectionBudget)
                        : Error(row, column,
                            "Decimal128 value exceeds XLSX numeric precision.");
                }
                if (hasDouble) return FiniteNumber(row, column, doubleValue, hasFormula, formulaIdentifier, formulas, options, projectionBudget);
                return hasFormula ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget) : Error(row, column, "Number cell has no value field.");
            case 3:
                if (hasString) {
                    if (strings.TryGetValue(stringIdentifier, out string? text)) {
                        projectionBudget.AddTextCharacters(text.Length);
                        return hasFormula
                            ? Formula(row, column, formulaIdentifier, formulas, options,
                                projectionBudget, text, IWorkCellKind.Text)
                            : new IWorkTableCell(row, column, IWorkCellKind.Text, text);
                    }
                    return hasFormula
                        ? Formula(row, column, formulaIdentifier, formulas, options,
                            projectionBudget, cachedValueIsComplete: false)
                        : Error(row, column, $"Unresolved shared string {stringIdentifier}.");
                }
                return hasFormula
                    ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget)
                    : Error(row, column, "Text cell has no shared-string value field.");
            case 5:
                if (!hasDate) return Error(row, column, "Date cell has no date value field.");
                if (!IsFinite(dateValue)) return Error(row, column, "Date cell has a non-finite value.");
                if (!TryReadDateTime(dateValue, out DateTime value)) {
                    return Error(row, column, "Date cell is outside the supported DateTime range.");
                }
                return hasFormula
                    ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget, value, IWorkCellKind.DateTime)
                    : new IWorkTableCell(row, column, IWorkCellKind.DateTime, value);
            case 6:
                if (!hasDouble) return Error(row, column, "Boolean cell has no value field.");
                if (!IsFinite(doubleValue)) return Error(row, column, "Boolean cell has a non-finite value.");
                if (doubleValue is not 0d and not 1d) {
                    return Error(row, column, "Boolean cell value is not 0 or 1.");
                }
                bool booleanValue = doubleValue == 1d;
                return hasFormula
                    ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget, booleanValue, IWorkCellKind.Boolean)
                    : new IWorkTableCell(row, column, IWorkCellKind.Boolean, booleanValue);
            case 7:
                if (!hasDouble) return Error(row, column, "Duration cell has no value field.");
                if (!IsFinite(doubleValue)) return Error(row, column, "Duration cell has a non-finite value.");
                return hasFormula
                    ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget, doubleValue, IWorkCellKind.Duration)
                    : new IWorkTableCell(row, column, IWorkCellKind.Duration, doubleValue);
            case 8:
                return hasFormula
                    ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget, "#ERROR", IWorkCellKind.Error)
                    : new IWorkTableCell(row, column, IWorkCellKind.Error, null, error: "#ERROR");
            case 9:
                if (hasRichString) {
                    if (hasFormula) formulaRichStringIdentifiers.Add(richStringIdentifier);
                    else nonFormulaRichStringIdentifiers.Add(richStringIdentifier);
                    if (richStrings.TryRead(richStringIdentifier, out IWorkTextContent? richText) && richText != null) {
                        string text = richText.PlainText;
                        projectionBudget.AddTextContentUse(richText, includeCharacters: true);
                        return hasFormula
                            ? Formula(row, column, formulaIdentifier, formulas, options,
                                projectionBudget, text, IWorkCellKind.Text,
                                cachedValueIsComplete: richText.IsTextComplete,
                                richText: richText)
                            : new IWorkTableCell(row, column, IWorkCellKind.Text, text,
                                richText: richText);
                    }
                    return hasFormula
                        ? Formula(row, column, formulaIdentifier, formulas, options,
                            projectionBudget, cachedValueIsComplete: false)
                        : Error(row, column, $"Unresolved rich text {richStringIdentifier}.");
                }
                return hasFormula ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget)
                    : new IWorkTableCell(row, column, IWorkCellKind.Text, string.Empty);
            default:
                return Error(row, column, $"Unknown cell type {type}.");
        }
    }

    private static IWorkTableCell Formula(int row, int column, uint formulaIdentifier,
        IReadOnlyDictionary<uint, IWorkWireMessage> formulas, IWorkReadOptions options,
        IWorkProjectionBudget projectionBudget,
        object? cachedValue = null,
        IWorkCellKind? cachedValueKind = null,
        bool cachedValueIsComplete = true,
        IWorkTextContent? richText = null) {
        IWorkFormulaResult result;
        if (formulas.TryGetValue(formulaIdentifier, out IWorkWireMessage? formula)) {
            projectionBudget.AddFormulaRenderingOperations(
                IWorkFormulaReader.MeasureRenderingOperations(formula,
                    options.MaximumFormulaNodes));
            result = IWorkFormulaReader.Render(formula, row - 1, column - 1,
                options.MaximumFormulaNodes, options.MaximumFormulaCharacters);
        } else {
            result = new IWorkFormulaResult("=?", false);
        }
        string formulaText = result.Text.Length == 0 ? "=?" : result.Text;
        projectionBudget.AddTextCharacters(formulaText.Length);
        projectionBudget.AddTextItem();
        return new IWorkTableCell(row, column, IWorkCellKind.Formula, cachedValue,
            formula: formulaText, valueKind: cachedValueKind,
            formulaIsComplete: result.IsComplete,
            richText: richText,
            cachedValueIsComplete: cachedValueIsComplete);
    }

    private static IWorkTableCell FiniteNumber(int row, int column, double value, bool hasFormula,
        uint formulaIdentifier, IReadOnlyDictionary<uint, IWorkWireMessage> formulas,
        IWorkReadOptions options, IWorkProjectionBudget projectionBudget) =>
        IsFinite(value)
            ? hasFormula
                ? Formula(row, column, formulaIdentifier, formulas, options, projectionBudget, value, IWorkCellKind.Number)
                : new IWorkTableCell(row, column, IWorkCellKind.Number, value)
            : Error(row, column, "Number cell has a non-finite value.");

    private static IWorkTableCell Error(int row, int column, string message) =>
        new(row, column, IWorkCellKind.Error, null, error: message, hasDecodeError: true);

    private static void MarkCellStorageUnsupported(IWorkArchiveRecord tile,
        ICollection<IWorkDiagnostic> diagnostics, ref bool supportsEditableReconstruction) {
        supportsEditableReconstruction = false;
        if (!diagnostics.Any(diagnostic => diagnostic.Code == "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED")) {
            diagnostics.Add(new IWorkDiagnostic(IWorkDiagnosticSeverity.Warning,
                "IWORK_TABLE_CELL_STORAGE_UNSUPPORTED",
                "An iWork table row declares malformed or incomplete modern cell storage; editable reconstruction is incomplete.",
                tile.EntryPath, tile.Identifier));
        }
    }

    private static double ReadDouble(byte[] buffer, int offset) =>
        BitConverter.Int64BitsToDouble(unchecked((long)IWorkProtobuf.ReadUInt64(buffer, offset)));

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static bool TryReadDateTime(double seconds, out DateTime value) {
        long epochTicks = new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc).Ticks;
        double deltaTicks = seconds * TimeSpan.TicksPerSecond;
        double roundedDeltaTicks = Math.Round(deltaTicks,
            MidpointRounding.AwayFromZero);
        if (!IsFinite(roundedDeltaTicks)
            || deltaTicks != roundedDeltaTicks
            || roundedDeltaTicks < -epochTicks
            || roundedDeltaTicks > DateTime.MaxValue.Ticks - epochTicks) {
            value = default;
            return false;
        }
        long absoluteTicks = epochTicks + (long)roundedDeltaTicks;
        if (absoluteTicks < DateTime.MinValue.Ticks
            || absoluteTicks > DateTime.MaxValue.Ticks) {
            value = default;
            return false;
        }
        value = new DateTime(absoluteTicks, DateTimeKind.Utc);
        return true;
    }

    private static double? ReadDecimal128(byte[] buffer, int offset) {
        if ((buffer[offset + 15] & 0x78) == 0x78) return null;
        int exponent = (((buffer[offset + 15] & 0x7f) << 7) | (buffer[offset + 14] >> 1)) - 0x1820;
        BigInteger coefficient = BigInteger.Zero;
        for (int index = 13; index >= 0; index--) {
            coefficient = coefficient * 256 + buffer[offset + index];
        }
        if ((buffer[offset + 14] & 1) != 0) coefficient += BigInteger.One << 112;
        if (coefficient.IsZero) return 0d;
        while (coefficient % 10 == 0) {
            coefficient /= 10;
            exponent++;
        }
        if (coefficient.ToString(CultureInfo.InvariantCulture).Length > 15) return null;
        string text = coefficient.ToString(CultureInfo.InvariantCulture)
            + "E" + exponent.ToString(CultureInfo.InvariantCulture);
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture,
                out double value) || !IsFinite(value) || value == 0d) return null;
        return (buffer[offset + 15] & 0x80) != 0 ? -value : value;
    }
}
