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
            while (offset < endOffset) {
                int tokenOffset = offset - expressionOffset;
                byte token = formulaPayload[offset++];
                ushort? functionId = null;
                string? functionName = null;
                byte? attribute = null;
                string? attributeName = null;

                if (IsFixedFunctionToken(token)) {
                    if (offset + 2 > endOffset) {
                        AddRecord(records, context, sheetName, cellReference, recordOffset, recordType, token, tokenOffset);
                        return;
                    }

                    functionId = BiffRecordReader.ReadUInt16(formulaPayload, offset);
                    functionName = GetFunctionName(functionId.Value);
                    offset += 2;
                } else if (IsVariableFunctionToken(token)) {
                    if (offset + 3 > endOffset) {
                        AddRecord(records, context, sheetName, cellReference, recordOffset, recordType, token, tokenOffset);
                        return;
                    }

                    offset++;
                    ushort functionBits = BiffRecordReader.ReadUInt16(formulaPayload, offset);
                    functionId = (ushort)(functionBits & 0x7fff);
                    functionName = (functionBits & 0x8000) != 0
                        ? $"Cetab:0x{functionId.Value:X4}"
                        : GetFunctionName(functionId.Value);
                    offset += 2;
                } else if (token == 0x19) {
                    if (offset >= endOffset) {
                        AddRecord(records, context, sheetName, cellReference, recordOffset, recordType, token, tokenOffset);
                        return;
                    }

                    attribute = formulaPayload[offset++];
                    attributeName = GetAttributeName(attribute.Value);
                    if (!TrySkipAttributePayload(formulaPayload, ref offset, endOffset, attribute.Value)) {
                        AddRecord(records, context, sheetName, cellReference, recordOffset, recordType, token, tokenOffset, functionId, functionName, attribute, attributeName);
                        return;
                    }
                } else if (!TrySkipTokenPayload(formulaPayload, ref offset, endOffset, token)) {
                    AddRecord(records, context, sheetName, cellReference, recordOffset, recordType, token, tokenOffset);
                    return;
                }

                AddRecord(records, context, sheetName, cellReference, recordOffset, recordType, token, tokenOffset, functionId, functionName, attribute, attributeName);
            }
        }

        private static void AddRecord(
            IList<LegacyXlsFormulaTokenRecord> records,
            string context,
            string? sheetName,
            string? cellReference,
            int recordOffset,
            ushort recordType,
            byte token,
            int tokenOffset,
            ushort? functionId = null,
            string? functionName = null,
            byte? attribute = null,
            string? attributeName = null) {
            records.Add(new LegacyXlsFormulaTokenRecord(
                context,
                sheetName,
                cellReference,
                recordOffset,
                recordType,
                token,
                BiffFormulaTokenInfo.GetTokenName(token),
                tokenOffset,
                functionId,
                functionName,
                attribute,
                attributeName));
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
