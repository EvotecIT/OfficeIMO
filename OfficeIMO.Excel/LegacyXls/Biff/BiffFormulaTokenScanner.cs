using System.Globalization;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Scans BIFF parsed-formula token streams for import-report diagnostics without evaluating the formula.
    /// </summary>
    internal static class BiffFormulaTokenScanner {
        internal static void ScanLengthPrefixed(
            byte[] formulaPayload,
            int parsedFormulaOffset,
            string context,
            string? sheetName,
            string? cellReference,
            int recordOffset,
            ushort recordType,
            IList<LegacyXlsFormulaTokenRecord> records) {
            if (parsedFormulaOffset + 2 > formulaPayload.Length) {
                return;
            }

            ushort expressionLength = BiffRecordReader.ReadUInt16(formulaPayload, parsedFormulaOffset);
            int expressionOffset = parsedFormulaOffset + 2;
            if (expressionLength == 0 || expressionOffset + expressionLength > formulaPayload.Length) {
                return;
            }

            ScanTokens(
                formulaPayload,
                expressionOffset,
                expressionOffset + expressionLength,
                context,
                sheetName,
                cellReference,
                recordOffset,
                recordType,
                records);
        }

        internal static void ScanTokens(
            byte[] formulaTokens,
            string context,
            string? sheetName,
            string? cellReference,
            int recordOffset,
            ushort recordType,
            IList<LegacyXlsFormulaTokenRecord> records) {
            ScanTokens(
                formulaTokens,
                0,
                formulaTokens.Length,
                context,
                sheetName,
                cellReference,
                recordOffset,
                recordType,
                records);
        }

        private static void ScanTokens(
            byte[] formulaPayload,
            int expressionOffset,
            int endOffset,
            string context,
            string? sheetName,
            string? cellReference,
            int recordOffset,
            ushort recordType,
            IList<LegacyXlsFormulaTokenRecord> records) {
            int offset = expressionOffset;
            int sequenceIndex = 0;
            while (offset < endOffset) {
                int tokenOffset = offset - expressionOffset;
                byte token = formulaPayload[offset++];
                int operandOffset = offset;
                ushort? functionId = null;
                string? functionName = null;
                byte? functionParameterCount = null;
                bool? functionIsCetab = null;
                byte? attribute = null;
                string? attributeName = null;

                if (IsFixedFunctionToken(token)) {
                    if (offset + 2 > endOffset) {
                        AddRecord(records, formulaPayload, operandOffset, context, sheetName, cellReference, recordOffset, recordType, token, tokenOffset, sequenceIndex++, endOffset - operandOffset);
                        return;
                    }

                    functionId = BiffRecordReader.ReadUInt16(formulaPayload, offset);
                    if (BiffFormulaFunctionMetadata.TryGetFixedFunctionMetadata(functionId.Value, out string? fixedFunctionName, out int fixedParameterCount)) {
                        functionName = fixedFunctionName;
                        if (fixedParameterCount >= 0 && fixedParameterCount <= byte.MaxValue) {
                            functionParameterCount = (byte)fixedParameterCount;
                        }
                    } else {
                        functionName = GetFunctionName(functionId.Value);
                    }

                    offset += 2;
                } else if (IsVariableFunctionToken(token)) {
                    if (offset + 3 > endOffset) {
                        AddRecord(records, formulaPayload, operandOffset, context, sheetName, cellReference, recordOffset, recordType, token, tokenOffset, sequenceIndex++, endOffset - operandOffset);
                        return;
                    }

                    functionParameterCount = formulaPayload[offset++];
                    ushort functionBits = BiffRecordReader.ReadUInt16(formulaPayload, offset);
                    functionId = (ushort)(functionBits & 0x7fff);
                    functionIsCetab = (functionBits & 0x8000) != 0;
                    functionName = functionIsCetab.Value
                        ? $"Cetab:0x{functionId.Value:X4}"
                        : GetFunctionName(functionId.Value);
                    offset += 2;
                } else if (token == 0x19) {
                    if (offset >= endOffset) {
                        AddRecord(records, formulaPayload, operandOffset, context, sheetName, cellReference, recordOffset, recordType, token, tokenOffset, sequenceIndex++, endOffset - operandOffset);
                        return;
                    }

                    attribute = formulaPayload[offset++];
                    attributeName = GetAttributeName(attribute.Value);
                    if (!TrySkipAttributePayload(formulaPayload, ref offset, endOffset, attribute.Value)) {
                        AddRecord(records, formulaPayload, operandOffset, context, sheetName, cellReference, recordOffset, recordType, token, tokenOffset, sequenceIndex++, endOffset - operandOffset, functionId, functionName, functionParameterCount, functionIsCetab, attribute, attributeName);
                        return;
                    }
                } else if (!TrySkipTokenPayload(formulaPayload, ref offset, endOffset, token)) {
                    AddRecord(records, formulaPayload, operandOffset, context, sheetName, cellReference, recordOffset, recordType, token, tokenOffset, sequenceIndex++, endOffset - operandOffset);
                    return;
                }

                AddRecord(records, formulaPayload, operandOffset, context, sheetName, cellReference, recordOffset, recordType, token, tokenOffset, sequenceIndex++, offset - operandOffset, functionId, functionName, functionParameterCount, functionIsCetab, attribute, attributeName);
            }
        }

        private static void AddRecord(
            IList<LegacyXlsFormulaTokenRecord> records,
            byte[] formulaPayload,
            int operandOffset,
            string context,
            string? sheetName,
            string? cellReference,
            int recordOffset,
            ushort recordType,
            byte token,
            int tokenOffset,
            int sequenceIndex,
            int operandByteCount,
            ushort? functionId = null,
            string? functionName = null,
            byte? functionParameterCount = null,
            bool? functionIsCetab = null,
            byte? attribute = null,
            string? attributeName = null) {
            DecodeTokenOperand(formulaPayload, operandOffset, token, operandByteCount, functionId, functionName, attributeName, out string? operandKind, out string? operandText);
            records.Add(new LegacyXlsFormulaTokenRecord(
                context,
                sheetName,
                cellReference,
                recordOffset,
                recordType,
                token,
                BiffFormulaTokenInfo.GetTokenName(token),
                tokenOffset,
                sequenceIndex,
                BiffFormulaTokenInfo.GetTokenClassName(token),
                operandByteCount,
                functionId,
                functionName,
                functionParameterCount,
                functionIsCetab,
                attribute,
                attributeName,
                operandKind,
                operandText));
        }

        private static void DecodeTokenOperand(
            byte[] payload,
            int offset,
            byte token,
            int operandByteCount,
            ushort? functionId,
            string? functionName,
            string? attributeName,
            out string? operandKind,
            out string? operandText) {
            operandKind = null;
            operandText = null;
            if (operandByteCount < 0 || offset + operandByteCount > payload.Length) {
                return;
            }

            if (IsFixedFunctionToken(token)) {
                operandKind = "FixedFunction";
                operandText = functionName ?? (functionId.HasValue ? $"Function:0x{functionId.Value:X4}" : null);
                return;
            }

            if (IsVariableFunctionToken(token)) {
                operandKind = "VariableFunction";
                operandText = functionName ?? (functionId.HasValue ? $"Function:0x{functionId.Value:X4}" : null);
                return;
            }

            switch (token) {
                case 0x01:
                    TryDecodeCellLocation(payload, offset, operandByteCount, "SharedOrArrayFormulaAnchor", out operandKind, out operandText);
                    break;
                case 0x02:
                    TryDecodeCellLocation(payload, offset, operandByteCount, "DataTableAnchor", out operandKind, out operandText);
                    break;
                case 0x17:
                    operandKind = "StringLiteral";
                    operandText = TryDecodeStringLiteral(payload, offset, operandByteCount);
                    break;
                case 0x19:
                    operandKind = "Attribute";
                    operandText = attributeName;
                    break;
                case 0x1c:
                    operandKind = "ErrorLiteral";
                    operandText = operandByteCount >= 1 ? BiffErrorValue.ToText(payload[offset]) : null;
                    break;
                case 0x1d:
                    operandKind = "BooleanLiteral";
                    operandText = operandByteCount >= 1 ? (payload[offset] == 0 ? "FALSE" : "TRUE") : null;
                    break;
                case 0x1e:
                    operandKind = "IntegerLiteral";
                    operandText = operandByteCount >= 2 ? BiffRecordReader.ReadUInt16(payload, offset).ToString(CultureInfo.InvariantCulture) : null;
                    break;
                case 0x1f:
                    operandKind = "NumberLiteral";
                    operandText = operandByteCount >= 8 ? BiffRecordReader.ReadDouble(payload, offset).ToString("G15", CultureInfo.InvariantCulture) : null;
                    break;
                case 0x20:
                case 0x40:
                case 0x60:
                    operandKind = "ArrayLiteral";
                    break;
                case 0x23:
                case 0x43:
                case 0x63:
                    operandKind = "DefinedName";
                    operandText = operandByteCount >= 4 ? $"NameIndex:{BiffRecordReader.ReadUInt32(payload, offset).ToString(CultureInfo.InvariantCulture)}" : null;
                    break;
                case 0x24:
                case 0x44:
                case 0x64:
                    operandKind = "CellReference";
                    operandText = operandByteCount >= 4 ? BiffFormulaReferenceFormatter.FormatCellReference(BiffRecordReader.ReadUInt16(payload, offset), BiffRecordReader.ReadUInt16(payload, offset + 2)) : null;
                    break;
                case 0x25:
                case 0x45:
                case 0x65:
                    operandKind = "AreaReference";
                    operandText = operandByteCount >= 8 ? FormatAreaReference(payload, offset) : null;
                    break;
                case 0x26:
                case 0x46:
                case 0x66:
                    operandKind = "MemoryArea";
                    break;
                case 0x27:
                case 0x47:
                case 0x67:
                    operandKind = "MemoryError";
                    break;
                case 0x28:
                case 0x48:
                case 0x68:
                    operandKind = "MemoryNoMemory";
                    break;
                case 0x29:
                case 0x49:
                case 0x69:
                    operandKind = "MemoryFunction";
                    break;
                case 0x2a:
                case 0x4a:
                case 0x6a:
                    operandKind = "DeletedCellReference";
                    break;
                case 0x2b:
                case 0x4b:
                case 0x6b:
                    operandKind = "DeletedAreaReference";
                    break;
                case 0x2c:
                case 0x4c:
                case 0x6c:
                    operandKind = "RelativeCellReference";
                    operandText = operandByteCount >= 4 ? FormatRawRelativeReference(payload, offset) : null;
                    break;
                case 0x2d:
                case 0x4d:
                case 0x6d:
                    operandKind = "RelativeAreaReference";
                    operandText = operandByteCount >= 8 ? FormatRawRelativeAreaReference(payload, offset) : null;
                    break;
                case 0x39:
                case 0x59:
                case 0x79:
                    operandKind = "ExternalName";
                    operandText = operandByteCount >= 6
                        ? $"ExternSheet:{BiffRecordReader.ReadUInt16(payload, offset).ToString(CultureInfo.InvariantCulture)};NameIndex:{BiffRecordReader.ReadUInt32(payload, offset + 2).ToString(CultureInfo.InvariantCulture)}"
                        : null;
                    break;
                case 0x3a:
                case 0x5a:
                case 0x7a:
                    operandKind = "ExternalCellReference";
                    operandText = operandByteCount >= 6 ? FormatExternalCellReference(payload, offset) : null;
                    break;
                case 0x3b:
                case 0x5b:
                case 0x7b:
                    operandKind = "ExternalAreaReference";
                    operandText = operandByteCount >= 10 ? FormatExternalAreaReference(payload, offset) : null;
                    break;
                case 0x3c:
                case 0x5c:
                case 0x7c:
                    operandKind = "DeletedExternalCellReference";
                    operandText = operandByteCount >= 6 ? $"ExternSheet:{BiffRecordReader.ReadUInt16(payload, offset).ToString(CultureInfo.InvariantCulture)}" : null;
                    break;
                case 0x3d:
                case 0x5d:
                case 0x7d:
                    operandKind = "DeletedExternalAreaReference";
                    operandText = operandByteCount >= 10 ? $"ExternSheet:{BiffRecordReader.ReadUInt16(payload, offset).ToString(CultureInfo.InvariantCulture)}" : null;
                    break;
            }
        }

        private static bool TrySkipTokenPayload(byte[] payload, ref int offset, int endOffset, byte token) {
            switch (token) {
                case 0x01:
                case 0x02:
                case 0x23:
                case 0x43:
                case 0x63:
                case 0x24:
                case 0x44:
                case 0x64:
                case 0x2a:
                case 0x4a:
                case 0x6a:
                case 0x2c:
                case 0x4c:
                case 0x6c:
                    return TrySkipBytes(ref offset, endOffset, 4);
                case 0x17:
                    return TrySkipStringLiteral(payload, ref offset, endOffset);
                case 0x1c:
                case 0x1d:
                    return TrySkipBytes(ref offset, endOffset, 1);
                case 0x1e:
                case 0x29:
                case 0x49:
                case 0x69:
                    return TrySkipBytes(ref offset, endOffset, 2);
                case 0x1f:
                    return TrySkipBytes(ref offset, endOffset, 8);
                case 0x20:
                case 0x40:
                case 0x60:
                    return TrySkipBytes(ref offset, endOffset, 7);
                case 0x26:
                case 0x46:
                case 0x66:
                case 0x27:
                case 0x47:
                case 0x67:
                case 0x28:
                case 0x48:
                case 0x68:
                    return TrySkipBytes(ref offset, endOffset, 6);
                case 0x25:
                case 0x45:
                case 0x65:
                case 0x2b:
                case 0x4b:
                case 0x6b:
                case 0x2d:
                case 0x4d:
                case 0x6d:
                    return TrySkipBytes(ref offset, endOffset, 8);
                case 0x39:
                case 0x59:
                case 0x79:
                case 0x3a:
                case 0x5a:
                case 0x7a:
                case 0x3c:
                case 0x5c:
                case 0x7c:
                    return TrySkipBytes(ref offset, endOffset, 6);
                case 0x3b:
                case 0x5b:
                case 0x7b:
                case 0x3d:
                case 0x5d:
                case 0x7d:
                    return TrySkipBytes(ref offset, endOffset, 10);
                default:
                    return IsNoPayloadToken(token);
            }
        }

        private static bool TrySkipAttributePayload(byte[] payload, ref int offset, int endOffset, byte attribute) {
            if (attribute == 0x04) {
                if (offset + 2 > endOffset) {
                    return false;
                }

                ushort choiceCount = BiffRecordReader.ReadUInt16(payload, offset);
                offset += 2;
                int offsetBytes = checked((choiceCount + 1) * 2);
                return TrySkipBytes(ref offset, endOffset, offsetBytes);
            }

            return TrySkipBytes(ref offset, endOffset, 2);
        }

        private static void TryDecodeCellLocation(byte[] payload, int offset, int operandByteCount, string kind, out string? operandKind, out string? operandText) {
            operandKind = operandByteCount >= 4 ? kind : null;
            operandText = operandByteCount >= 4
                ? A1.ColumnIndexToLetters(BiffRecordReader.ReadUInt16(payload, offset + 2) + 1) + (BiffRecordReader.ReadUInt16(payload, offset) + 1).ToString(CultureInfo.InvariantCulture)
                : null;
        }

        private static string? TryDecodeStringLiteral(byte[] payload, int offset, int operandByteCount) {
            try {
                int stringOffset = offset;
                string value = BiffStringReader.ReadShortUnicodeString(payload, ref stringOffset);
                return stringOffset <= offset + operandByteCount ? value : null;
            } catch (InvalidDataException) {
                return null;
            } catch (OverflowException) {
                return null;
            }
        }

        private static string FormatAreaReference(byte[] payload, int offset) {
            string start = BiffFormulaReferenceFormatter.FormatCellReference(BiffRecordReader.ReadUInt16(payload, offset), BiffRecordReader.ReadUInt16(payload, offset + 4));
            string end = BiffFormulaReferenceFormatter.FormatCellReference(BiffRecordReader.ReadUInt16(payload, offset + 2), BiffRecordReader.ReadUInt16(payload, offset + 6));
            return start == end ? start : start + ":" + end;
        }

        private static string FormatRawRelativeReference(byte[] payload, int offset) {
            return $"Row:{unchecked((short)BiffRecordReader.ReadUInt16(payload, offset)).ToString(CultureInfo.InvariantCulture)};ColumnBits:0x{BiffRecordReader.ReadUInt16(payload, offset + 2):X4}";
        }

        private static string FormatRawRelativeAreaReference(byte[] payload, int offset) {
            return $"Rows:{unchecked((short)BiffRecordReader.ReadUInt16(payload, offset)).ToString(CultureInfo.InvariantCulture)}:{unchecked((short)BiffRecordReader.ReadUInt16(payload, offset + 2)).ToString(CultureInfo.InvariantCulture)};ColumnBits:0x{BiffRecordReader.ReadUInt16(payload, offset + 4):X4}:0x{BiffRecordReader.ReadUInt16(payload, offset + 6):X4}";
        }

        private static string FormatExternalCellReference(byte[] payload, int offset) {
            return $"ExternSheet:{BiffRecordReader.ReadUInt16(payload, offset).ToString(CultureInfo.InvariantCulture)};Reference:{BiffFormulaReferenceFormatter.FormatCellReference(BiffRecordReader.ReadUInt16(payload, offset + 2), BiffRecordReader.ReadUInt16(payload, offset + 4))}";
        }

        private static string FormatExternalAreaReference(byte[] payload, int offset) {
            return $"ExternSheet:{BiffRecordReader.ReadUInt16(payload, offset).ToString(CultureInfo.InvariantCulture)};Reference:{FormatAreaReference(payload, offset + 2)}";
        }

        private static bool TrySkipStringLiteral(byte[] payload, ref int offset, int endOffset) {
            try {
                int stringOffset = offset;
                BiffStringReader.ReadShortUnicodeString(payload, ref stringOffset);
                if (stringOffset > endOffset) {
                    return false;
                }

                offset = stringOffset;
                return true;
            } catch (InvalidDataException) {
                return false;
            } catch (OverflowException) {
                return false;
            }
        }

        private static bool TrySkipBytes(ref int offset, int endOffset, int bytes) {
            if (offset + bytes > endOffset) {
                return false;
            }

            offset += bytes;
            return true;
        }

        private static bool IsNoPayloadToken(byte token) {
            return token >= 0x03 && token <= 0x16;
        }

        private static bool IsFixedFunctionToken(byte token) {
            return token == 0x21 || token == 0x41 || token == 0x61;
        }

        private static bool IsVariableFunctionToken(byte token) {
            return token == 0x22 || token == 0x42 || token == 0x62;
        }

        private static string? GetFunctionName(ushort functionId) {
            return BiffFormulaFunctionMetadata.TryGetKnownFunctionName(functionId, out string? functionName)
                ? functionName
                : null;
        }

        private static string GetAttributeName(byte attribute) {
            return attribute switch {
                0x01 => "Volatile",
                0x02 => "If",
                0x04 => "Choose",
                0x08 => "Goto",
                0x10 => "Sum",
                0x40 => "Space",
                0x41 => "SpaceVolatile",
                _ => $"Attribute:0x{attribute:X2}"
            };
        }
    }
}
