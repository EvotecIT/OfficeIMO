using System.Globalization;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    /// <summary>
    /// Decodes supported BIFF formula token streams into Open XML formula text.
    /// Unsupported tokens return false so callers can preserve the cached result only.
    /// </summary>
    internal static class BiffFormulaTextReader {
        internal static bool TryRead(byte[] formulaPayload, int parsedFormulaOffset, out string? formulaText) {
            return TryRead(
                formulaPayload,
                parsedFormulaOffset,
                formulaRow: 0,
                formulaColumn: 0,
                Array.Empty<BiffExternSheetReference>(),
                Array.Empty<LegacyXlsExternalReference>(),
                Array.Empty<string>(),
                Array.Empty<string?>(),
                out formulaText);
        }

        internal static bool TryRead(
            byte[] formulaPayload,
            int parsedFormulaOffset,
            int formulaRow,
            int formulaColumn,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            out string? formulaText,
            out BiffFormulaReadFailure? failure) {
            return TryRead(
                formulaPayload,
                parsedFormulaOffset,
                formulaRow,
                formulaColumn,
                externSheets,
                Array.Empty<LegacyXlsExternalReference>(),
                sheetNames,
                definedNames,
                out formulaText,
                out failure);
        }

        internal static bool TryRead(
            byte[] formulaPayload,
            int parsedFormulaOffset,
            int formulaRow,
            int formulaColumn,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            out string? formulaText,
            out BiffFormulaReadFailure? failure) {
            if (TryRead(formulaPayload, parsedFormulaOffset, formulaRow, formulaColumn, externSheets, externalReferences, sheetNames, definedNames, out formulaText)) {
                failure = null;
                return true;
            }

            failure = DescribeFailure(formulaPayload, parsedFormulaOffset, externSheets, externalReferences, sheetNames, definedNames);
            return false;
        }

        internal static bool TryRead(
            byte[] formulaPayload,
            int parsedFormulaOffset,
            int formulaRow,
            int formulaColumn,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            out string? formulaText) {
            return TryRead(
                formulaPayload,
                parsedFormulaOffset,
                formulaRow,
                formulaColumn,
                externSheets,
                Array.Empty<LegacyXlsExternalReference>(),
                sheetNames,
                definedNames,
                out formulaText);
        }

        internal static bool TryRead(
            byte[] formulaPayload,
            int parsedFormulaOffset,
            int formulaRow,
            int formulaColumn,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames,
            out string? formulaText) {
            formulaText = null;
            if (parsedFormulaOffset + 2 > formulaPayload.Length) {
                return false;
            }

            ushort expressionLength = BiffRecordReader.ReadUInt16(formulaPayload, parsedFormulaOffset);
            int expressionOffset = parsedFormulaOffset + 2;
            if (expressionLength == 0 || expressionOffset + expressionLength > formulaPayload.Length) {
                return false;
            }

            var stack = new Stack<FormulaExpression>();
            int offset = expressionOffset;
            int endOffset = expressionOffset + expressionLength;
            int extraOffset = endOffset;
            while (offset < endOffset) {
                byte token = formulaPayload[offset++];
                switch (token) {
                    case 0x03:
                        if (!ApplyBinaryOperator(stack, "+", 1)) return false;
                        break;
                    case 0x04:
                        if (!ApplyBinaryOperator(stack, "-", 1)) return false;
                        break;
                    case 0x05:
                        if (!ApplyBinaryOperator(stack, "*", 2)) return false;
                        break;
                    case 0x06:
                        if (!ApplyBinaryOperator(stack, "/", 2)) return false;
                        break;
                    case 0x07:
                        if (!ApplyBinaryOperator(stack, "^", 3)) return false;
                        break;
                    case 0x08:
                        if (!ApplyBinaryOperator(stack, "&", 0)) return false;
                        break;
                    case 0x09:
                        if (!ApplyBinaryOperator(stack, "<", 0)) return false;
                        break;
                    case 0x0a:
                        if (!ApplyBinaryOperator(stack, "<=", 0)) return false;
                        break;
                    case 0x0b:
                        if (!ApplyBinaryOperator(stack, "=", 0)) return false;
                        break;
                    case 0x0c:
                        if (!ApplyBinaryOperator(stack, ">=", 0)) return false;
                        break;
                    case 0x0d:
                        if (!ApplyBinaryOperator(stack, ">", 0)) return false;
                        break;
                    case 0x0e:
                        if (!ApplyBinaryOperator(stack, "<>", 0)) return false;
                        break;
                    case 0x0f:
                        if (!ApplyReferenceOperator(stack, " ", groupResult: true)) return false;
                        break;
                    case 0x10:
                        if (!ApplyReferenceOperator(stack, ",", groupResult: true)) return false;
                        break;
                    case 0x11:
                        if (!ApplyReferenceOperator(stack, ":", groupResult: false)) return false;
                        break;
                    case 0x12:
                        if (!ApplyUnaryOperator(stack, "+")) return false;
                        break;
                    case 0x13:
                        if (!ApplyUnaryOperator(stack, "-")) return false;
                        break;
                    case 0x14:
                        if (!ApplyPercentOperator(stack)) return false;
                        break;
                    case 0x15:
                        if (!ApplyParentheses(stack)) return false;
                        break;
                    case 0x16:
                        stack.Push(new FormulaExpression(string.Empty, 4));
                        break;
                    case 0x17:
                        if (!TryReadStringLiteral(formulaPayload, ref offset, endOffset, stack)) return false;
                        break;
                    case 0x19:
                        if (!TryReadAttributeToken(formulaPayload, ref offset, endOffset, stack)) return false;
                        break;
                    case 0x23:
                    case 0x43:
                    case 0x63:
                        if (offset + 4 > endOffset) return false;
                        if (!TryReadDefinedName(formulaPayload, offset, definedNames, out string? definedName)) return false;
                        stack.Push(new FormulaExpression(definedName!, 4));
                        offset += 4;
                        break;
                    case 0x39:
                    case 0x59:
                    case 0x79:
                        if (offset + 6 > endOffset) return false;
                        if (!BiffFormulaReferenceFormatter.TryReadExternalName(formulaPayload, offset, externSheets, externalReferences, out string? externalName)) return false;
                        stack.Push(new FormulaExpression(externalName!, 4));
                        offset += 6;
                        break;
                    case 0x21:
                    case 0x41:
                    case 0x61:
                        if (offset + 2 > endOffset) return false;
                        ushort fixedFunctionId = BiffRecordReader.ReadUInt16(formulaPayload, offset);
                        offset += 2;
                        if (!ApplyFixedFunction(stack, fixedFunctionId)) return false;
                        break;
                    case 0x22:
                    case 0x42:
                    case 0x62:
                        if (offset + 3 > endOffset) return false;
                        byte parameterCount = formulaPayload[offset++];
                        ushort functionBits = BiffRecordReader.ReadUInt16(formulaPayload, offset);
                        offset += 2;
                        if (!ApplyVariableFunction(stack, parameterCount, functionBits)) return false;
                        break;
                    case 0x1d:
                        if (offset >= endOffset) return false;
                        stack.Push(new FormulaExpression(formulaPayload[offset++] == 0 ? "FALSE" : "TRUE", 4));
                        break;
                    case 0x1c:
                        if (offset >= endOffset) return false;
                        stack.Push(new FormulaExpression(BiffErrorValue.ToText(formulaPayload[offset++]), 4));
                        break;
                    case 0x1e:
                        if (offset + 2 > endOffset) return false;
                        stack.Push(new FormulaExpression(BiffRecordReader.ReadUInt16(formulaPayload, offset).ToString(CultureInfo.InvariantCulture), 4));
                        offset += 2;
                        break;
                    case 0x1f:
                        if (offset + 8 > endOffset) return false;
                        stack.Push(new FormulaExpression(BiffRecordReader.ReadDouble(formulaPayload, offset).ToString("G15", CultureInfo.InvariantCulture), 4));
                        offset += 8;
                        break;
                    case 0x20:
                    case 0x40:
                    case 0x60:
                        if (offset + 7 > endOffset) return false;
                        offset += 7;
                        if (!BiffFormulaArrayConstantReader.TryRead(formulaPayload, ref extraOffset, out string? arrayText)) return false;
                        stack.Push(new FormulaExpression(arrayText!, 4));
                        break;
                    case 0x26:
                    case 0x46:
                    case 0x66:
                        if (!TrySkipMemAreaHeader(formulaPayload, ref offset, endOffset)) return false;
                        break;
                    case 0x29:
                    case 0x49:
                    case 0x69:
                        if (!TrySkipMemFuncHeader(formulaPayload, ref offset, endOffset)) return false;
                        break;
                    case 0x24:
                    case 0x44:
                    case 0x64:
                        if (offset + 4 > endOffset) return false;
                        stack.Push(new FormulaExpression(ReadCellReference(formulaPayload, offset), 4));
                        offset += 4;
                        break;
                    case 0x25:
                    case 0x45:
                    case 0x65:
                        if (offset + 8 > endOffset) return false;
                        stack.Push(new FormulaExpression(ReadAreaReference(formulaPayload, offset), 4));
                        offset += 8;
                        break;
                    case 0x2c:
                    case 0x4c:
                    case 0x6c:
                        if (offset + 4 > endOffset) return false;
                        stack.Push(new FormulaExpression(ReadRelativeCellReference(formulaPayload, offset, formulaRow, formulaColumn), 4));
                        offset += 4;
                        break;
                    case 0x2d:
                    case 0x4d:
                    case 0x6d:
                        if (offset + 8 > endOffset) return false;
                        stack.Push(new FormulaExpression(ReadRelativeAreaReference(formulaPayload, offset, formulaRow, formulaColumn), 4));
                        offset += 8;
                        break;
                    case 0x2a:
                    case 0x4a:
                    case 0x6a:
                        if (offset + 4 > endOffset) return false;
                        stack.Push(new FormulaExpression("#REF!", 4));
                        offset += 4;
                        break;
                    case 0x2b:
                    case 0x4b:
                    case 0x6b:
                        if (offset + 8 > endOffset) return false;
                        stack.Push(new FormulaExpression("#REF!", 4));
                        offset += 8;
                        break;
                    case 0x3a:
                    case 0x5a:
                    case 0x7a:
                        if (offset + 6 > endOffset) return false;
                        if (!BiffFormulaReferenceFormatter.TryRead3dReference(formulaPayload, offset, externSheets, externalReferences, sheetNames, out string? reference3d)) return false;
                        stack.Push(new FormulaExpression(reference3d!, 4));
                        offset += 6;
                        break;
                    case 0x3b:
                    case 0x5b:
                    case 0x7b:
                        if (offset + 10 > endOffset) return false;
                        if (!BiffFormulaReferenceFormatter.TryRead3dArea(formulaPayload, offset, externSheets, externalReferences, sheetNames, out string? area3d)) return false;
                        stack.Push(new FormulaExpression(area3d!, 4));
                        offset += 10;
                        break;
                    case 0x3c:
                    case 0x5c:
                    case 0x7c:
                        if (offset + 6 > endOffset) return false;
                        if (!BiffFormulaReferenceFormatter.TryRead3dInvalidReference(formulaPayload, offset, externSheets, externalReferences, sheetNames, out string? invalidReference3d)) return false;
                        stack.Push(new FormulaExpression(invalidReference3d!, 4));
                        offset += 6;
                        break;
                    case 0x3d:
                    case 0x5d:
                    case 0x7d:
                        if (offset + 10 > endOffset) return false;
                        if (!BiffFormulaReferenceFormatter.TryRead3dInvalidReference(formulaPayload, offset, externSheets, externalReferences, sheetNames, out string? invalidArea3d)) return false;
                        stack.Push(new FormulaExpression(invalidArea3d!, 4));
                        offset += 10;
                        break;
                    default:
                        return false;
                }
            }

            if (stack.Count != 1) {
                return false;
            }

            formulaText = stack.Pop().Text;
            return !string.IsNullOrWhiteSpace(formulaText);
        }

        private static BiffFormulaReadFailure DescribeFailure(
            byte[] formulaPayload,
            int parsedFormulaOffset,
            IReadOnlyList<BiffExternSheetReference> externSheets,
            IReadOnlyList<LegacyXlsExternalReference> externalReferences,
            IReadOnlyList<string> sheetNames,
            IReadOnlyList<string?> definedNames) {
            if (parsedFormulaOffset + 2 > formulaPayload.Length) {
                return BiffFormulaReadFailure.InvalidPayload("Formula payload ended before the parsed-expression length.");
            }

            ushort expressionLength = BiffRecordReader.ReadUInt16(formulaPayload, parsedFormulaOffset);
            int expressionOffset = parsedFormulaOffset + 2;
            if (expressionLength == 0) {
                return BiffFormulaReadFailure.InvalidPayload("Formula payload did not contain parsed-expression tokens.");
            }

            if (expressionOffset + expressionLength > formulaPayload.Length) {
                return BiffFormulaReadFailure.InvalidPayload("Formula payload ended before all parsed-expression tokens could be read.");
            }

            int offset = expressionOffset;
            int endOffset = expressionOffset + expressionLength;
            int extraOffset = endOffset;
            int stackDepth = 0;
            while (offset < endOffset) {
                int tokenOffset = offset - expressionOffset;
                byte token = formulaPayload[offset++];
                switch (token) {
                    case 0x03:
                    case 0x04:
                    case 0x05:
                    case 0x06:
                    case 0x07:
                    case 0x08:
                    case 0x09:
                    case 0x0a:
                    case 0x0b:
                    case 0x0c:
                    case 0x0d:
                    case 0x0e:
                    case 0x0f:
                    case 0x10:
                    case 0x11:
                        if (!TryApplyStackOperator(ref stackDepth, 2, 1)) {
                            return BiffFormulaReadFailure.Stack("Formula operator did not have enough operands on the expression stack.", token, tokenOffset);
                        }

                        break;
                    case 0x12:
                    case 0x13:
                    case 0x14:
                    case 0x15:
                        if (!TryApplyStackOperator(ref stackDepth, 1, 1)) {
                            return BiffFormulaReadFailure.Stack("Formula unary operator did not have an operand on the expression stack.", token, tokenOffset);
                        }

                        break;
                    case 0x16:
                        stackDepth++;
                        break;
                    case 0x17:
                        if (!TrySkipStringLiteral(formulaPayload, ref offset, endOffset)) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula string literal token could not be read.", token, tokenOffset);
                        }

                        stackDepth++;
                        break;
                    case 0x19:
                        if (!TrySkipAttributeToken(formulaPayload, ref offset, endOffset, out byte unsupportedAttribute)) {
                            return BiffFormulaReadFailure.UnsupportedAttribute(unsupportedAttribute, token, tokenOffset);
                        }

                        break;
                    case 0x1c:
                    case 0x1d:
                        if (!TrySkipBytes(ref offset, endOffset, 1)) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula error or Boolean token ended early.", token, tokenOffset);
                        }

                        stackDepth++;
                        break;
                    case 0x1e:
                        if (!TrySkipBytes(ref offset, endOffset, 2)) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula integer token ended early.", token, tokenOffset);
                        }

                        stackDepth++;
                        break;
                    case 0x1f:
                        if (!TrySkipBytes(ref offset, endOffset, 8)) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula number token ended early.", token, tokenOffset);
                        }

                        stackDepth++;
                        break;
                    case 0x20:
                    case 0x40:
                    case 0x60:
                        if (!TrySkipBytes(ref offset, endOffset, 7)) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula array token ended early.", token, tokenOffset);
                        }

                        if (!BiffFormulaArrayConstantReader.TryRead(formulaPayload, ref extraOffset, out _)) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula array constant payload could not be read.", token, tokenOffset);
                        }

                        stackDepth++;
                        break;
                    case 0x21:
                    case 0x41:
                    case 0x61:
                        if (offset + 2 > endOffset) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula fixed-function token ended early.", token, tokenOffset);
                        }

                        ushort fixedFunctionId = BiffRecordReader.ReadUInt16(formulaPayload, offset);
                        offset += 2;
                        if (!BiffFormulaFunctionMetadata.TryGetFixedFunctionMetadata(fixedFunctionId, out string? fixedFunctionName, out int fixedParameterCount)) {
                            BiffFormulaFunctionMetadata.TryGetKnownFunctionName(fixedFunctionId, out string? knownFunctionName);
                            return BiffFormulaReadFailure.UnsupportedFixedFunction(fixedFunctionId, knownFunctionName, token, tokenOffset);
                        }

                        if (!TryApplyFunction(ref stackDepth, fixedParameterCount)) {
                            return BiffFormulaReadFailure.FunctionStackUnderflow(fixedFunctionId, fixedFunctionName!, fixedParameterCount, stackDepth, token, tokenOffset);
                        }

                        break;
                    case 0x22:
                    case 0x42:
                    case 0x62:
                        if (offset + 3 > endOffset) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula variable-function token ended early.", token, tokenOffset);
                        }

                        byte parameterCount = formulaPayload[offset++];
                        ushort functionBits = BiffRecordReader.ReadUInt16(formulaPayload, offset);
                        offset += 2;
                        bool isCetabFunction = (functionBits & 0x8000) != 0;
                        ushort functionId = (ushort)(functionBits & 0x7fff);
                        if (isCetabFunction || !BiffFormulaFunctionMetadata.TryGetFunctionName(functionId, out string? functionName)) {
                            if (!isCetabFunction && functionId == 0x00ff && parameterCount > 0) {
                                if (!TryApplyFunction(ref stackDepth, parameterCount)) {
                                    return BiffFormulaReadFailure.FunctionStackUnderflow(functionId, "UserDefined", parameterCount, stackDepth, token, tokenOffset);
                                }

                                break;
                            }

                            BiffFormulaFunctionMetadata.TryGetKnownFunctionName(functionId, out string? knownFunctionName);
                            return BiffFormulaReadFailure.UnsupportedVariableFunction(functionId, knownFunctionName, isCetabFunction, token, tokenOffset);
                        }

                        if (!BiffFormulaFunctionMetadata.IsSupportedVariableFunctionArgumentCount(functionId, parameterCount)) {
                            return BiffFormulaReadFailure.UnsupportedFunctionArguments(functionId, parameterCount, token, tokenOffset);
                        }

                        if (!TryApplyFunction(ref stackDepth, parameterCount)) {
                            return BiffFormulaReadFailure.FunctionStackUnderflow(functionId, functionName!, parameterCount, stackDepth, token, tokenOffset);
                        }

                        break;
                    case 0x23:
                    case 0x43:
                    case 0x63:
                        if (offset + 4 > endOffset) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula defined-name token ended early.", token, tokenOffset);
                        }

                        uint oneBasedNameIndex = BiffRecordReader.ReadUInt32(formulaPayload, offset);
                        offset += 4;
                        if (oneBasedNameIndex == 0
                            || oneBasedNameIndex > definedNames.Count
                            || oneBasedNameIndex > int.MaxValue
                            || string.IsNullOrWhiteSpace(definedNames[checked((int)oneBasedNameIndex) - 1])) {
                            return BiffFormulaReadFailure.DefinedName(oneBasedNameIndex, token, tokenOffset);
                        }

                        stackDepth++;
                        break;
                    case 0x39:
                    case 0x59:
                    case 0x79:
                        if (offset + 6 > endOffset) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula external defined-name token ended early.", token, tokenOffset);
                        }

                        if (!BiffFormulaReferenceFormatter.TryReadExternalName(formulaPayload, offset, externSheets, externalReferences, out _)) {
                            ushort externSheetIndex = BiffRecordReader.ReadUInt16(formulaPayload, offset);
                            uint oneBasedExternalNameIndex = BiffRecordReader.ReadUInt32(formulaPayload, offset + 2);
                            return BiffFormulaReadFailure.ExternalName(externSheetIndex, oneBasedExternalNameIndex, token, tokenOffset);
                        }

                        stackDepth++;
                        offset += 6;
                        break;
                    case 0x26:
                    case 0x46:
                    case 0x66:
                        if (!TrySkipMemAreaHeader(formulaPayload, ref offset, endOffset)) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula mem-area token ended early.", token, tokenOffset);
                        }

                        break;
                    case 0x29:
                    case 0x49:
                    case 0x69:
                        if (!TrySkipMemFuncHeader(formulaPayload, ref offset, endOffset)) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula mem-function token ended early.", token, tokenOffset);
                        }

                        break;
                    case 0x24:
                    case 0x44:
                    case 0x64:
                    case 0x2a:
                    case 0x4a:
                    case 0x6a:
                    case 0x2c:
                    case 0x4c:
                    case 0x6c:
                        if (!TrySkipBytes(ref offset, endOffset, 4)) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula cell-reference token ended early.", token, tokenOffset);
                        }

                        stackDepth++;
                        break;
                    case 0x25:
                    case 0x45:
                    case 0x65:
                    case 0x2b:
                    case 0x4b:
                    case 0x6b:
                    case 0x2d:
                    case 0x4d:
                    case 0x6d:
                        if (!TrySkipBytes(ref offset, endOffset, 8)) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula area-reference token ended early.", token, tokenOffset);
                        }

                        stackDepth++;
                        break;
                    case 0x3a:
                    case 0x5a:
                    case 0x7a:
                        if (offset + 6 > endOffset) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula 3D reference token ended early.", token, tokenOffset);
                        }

                        if (!BiffFormulaReferenceFormatter.TryRead3dReference(formulaPayload, offset, externSheets, externalReferences, sheetNames, out _)) {
                            return BiffFormulaReadFailure.Reference("Formula3dReference", "Formula 3D reference could not be resolved.", token, tokenOffset);
                        }

                        stackDepth++;
                        offset += 6;
                        break;
                    case 0x3b:
                    case 0x5b:
                    case 0x7b:
                        if (offset + 10 > endOffset) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula 3D area token ended early.", token, tokenOffset);
                        }

                        if (!BiffFormulaReferenceFormatter.TryRead3dArea(formulaPayload, offset, externSheets, externalReferences, sheetNames, out _)) {
                            return BiffFormulaReadFailure.Reference("Formula3dArea", "Formula 3D area could not be resolved.", token, tokenOffset);
                        }

                        stackDepth++;
                        offset += 10;
                        break;
                    case 0x3c:
                    case 0x5c:
                    case 0x7c:
                        if (offset + 6 > endOffset) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula invalid 3D reference token ended early.", token, tokenOffset);
                        }

                        if (!BiffFormulaReferenceFormatter.TryRead3dInvalidReference(formulaPayload, offset, externSheets, externalReferences, sheetNames, out _)) {
                            return BiffFormulaReadFailure.Reference("FormulaInvalid3dReference", "Formula invalid 3D reference could not be resolved.", token, tokenOffset);
                        }

                        stackDepth++;
                        offset += 6;
                        break;
                    case 0x3d:
                    case 0x5d:
                    case 0x7d:
                        if (offset + 10 > endOffset) {
                            return BiffFormulaReadFailure.InvalidPayload("Formula invalid 3D area token ended early.", token, tokenOffset);
                        }

                        if (!BiffFormulaReferenceFormatter.TryRead3dInvalidReference(formulaPayload, offset, externSheets, externalReferences, sheetNames, out _)) {
                            return BiffFormulaReadFailure.Reference("FormulaInvalid3dArea", "Formula invalid 3D area could not be resolved.", token, tokenOffset);
                        }

                        stackDepth++;
                        offset += 10;
                        break;
                    default:
                        return BiffFormulaReadFailure.UnsupportedToken(token, tokenOffset);
                }
            }

            return BiffFormulaReadFailure.Stack("Formula tokens are individually recognized but the expression stack could not be reduced.");
        }

        private static bool TryApplyStackOperator(ref int stackDepth, int requiredOperands, int producedOperands) {
            if (stackDepth < requiredOperands) {
                return false;
            }

            stackDepth = stackDepth - requiredOperands + producedOperands;
            return true;
        }

        private static bool TryApplyFunction(ref int stackDepth, int parameterCount) {
            return TryApplyStackOperator(ref stackDepth, parameterCount, 1);
        }

        private static bool TrySkipMemAreaHeader(byte[] formulaPayload, ref int offset, int endOffset) {
            if (offset + 6 > endOffset) {
                return false;
            }

            ushort expressionBytes = BiffRecordReader.ReadUInt16(formulaPayload, offset + 4);
            if (expressionBytes == 0 || offset + 6 + expressionBytes > endOffset) {
                return false;
            }

            offset += 6;
            return true;
        }

        private static bool TrySkipMemFuncHeader(byte[] formulaPayload, ref int offset, int endOffset) {
            if (offset + 2 > endOffset) {
                return false;
            }

            ushort expressionBytes = BiffRecordReader.ReadUInt16(formulaPayload, offset);
            if (expressionBytes == 0 || offset + 2 + expressionBytes > endOffset) {
                return false;
            }

            offset += 2;
            return true;
        }

        private static bool TrySkipStringLiteral(byte[] formulaPayload, ref int offset, int endOffset) {
            try {
                int stringOffset = offset;
                BiffStringReader.ReadShortUnicodeString(formulaPayload, ref stringOffset);
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

        private static bool TrySkipAttributeToken(byte[] formulaPayload, ref int offset, int endOffset, out byte unsupportedAttribute) {
            unsupportedAttribute = 0;
            if (offset >= endOffset) {
                return false;
            }

            byte attribute = formulaPayload[offset++];
            if (attribute == 0x04) {
                if (offset + 2 > endOffset) {
                    unsupportedAttribute = attribute;
                    return false;
                }

                ushort choiceCount = BiffRecordReader.ReadUInt16(formulaPayload, offset);
                offset += 2;
                if (choiceCount == 0 || choiceCount > 29) {
                    unsupportedAttribute = attribute;
                    return false;
                }

                int offsetBytes = checked((choiceCount + 1) * 2);
                if (offset + offsetBytes > endOffset) {
                    unsupportedAttribute = attribute;
                    return false;
                }

                offset += offsetBytes;
                return true;
            }

            if (offset + 2 > endOffset) {
                unsupportedAttribute = attribute;
                return false;
            }

            offset += 2;
            if (attribute == 0x01 || attribute == 0x02 || attribute == 0x08 || attribute == 0x10 || attribute == 0x40 || attribute == 0x41) {
                return true;
            }

            unsupportedAttribute = attribute;
            return false;
        }

        private static bool TrySkipBytes(ref int offset, int endOffset, int bytes) {
            if (offset + bytes > endOffset) {
                return false;
            }

            offset += bytes;
            return true;
        }

        private static bool TryReadDefinedName(byte[] formulaPayload, int offset, IReadOnlyList<string?> definedNames, out string? name) {
            name = null;
            uint oneBasedNameIndex = BiffRecordReader.ReadUInt32(formulaPayload, offset);
            if (oneBasedNameIndex == 0 || oneBasedNameIndex > definedNames.Count || oneBasedNameIndex > int.MaxValue) {
                return false;
            }

            name = definedNames[checked((int)oneBasedNameIndex) - 1];
            return !string.IsNullOrWhiteSpace(name);
        }

        private static bool TryReadStringLiteral(byte[] formulaPayload, ref int offset, int endOffset, Stack<FormulaExpression> stack) {
            try {
                int stringOffset = offset;
                string value = BiffStringReader.ReadShortUnicodeString(formulaPayload, ref stringOffset);
                if (stringOffset > endOffset) {
                    return false;
                }

                offset = stringOffset;
                stack.Push(new FormulaExpression(QuoteFormulaStringLiteral(value), 4));
                return true;
            } catch (InvalidDataException) {
                return false;
            } catch (OverflowException) {
                return false;
            }
        }

        private static string QuoteFormulaStringLiteral(string value) {
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private static bool TryReadAttributeToken(byte[] formulaPayload, ref int offset, int endOffset, Stack<FormulaExpression> stack) {
            if (offset >= endOffset) {
                return false;
            }

            byte attribute = formulaPayload[offset++];
            if (attribute == 0x04) {
                if (offset + 2 > endOffset) {
                    return false;
                }

                ushort choiceCount = BiffRecordReader.ReadUInt16(formulaPayload, offset);
                offset += 2;
                if (choiceCount == 0 || choiceCount > 29) {
                    return false;
                }

                int offsetBytes = checked((choiceCount + 1) * 2);
                if (offset + offsetBytes > endOffset) {
                    return false;
                }

                offset += offsetBytes;
                return true;
            }

            if (offset + 2 > endOffset) {
                return false;
            }

            offset += 2;
            if (attribute == 0x01 || attribute == 0x02 || attribute == 0x08 || attribute == 0x40 || attribute == 0x41) {
                return true;
            }

            return attribute == 0x10 && ApplySumAttribute(stack);
        }

        private static bool ApplyBinaryOperator(Stack<FormulaExpression> stack, string op, int precedence) {
            if (stack.Count < 2) {
                return false;
            }

            FormulaExpression right = stack.Pop();
            FormulaExpression left = stack.Pop();
            string leftText = left.Precedence < precedence ? "(" + left.Text + ")" : left.Text;
            string rightText = right.Precedence <= precedence ? "(" + right.Text + ")" : right.Text;
            stack.Push(new FormulaExpression(leftText + op + rightText, precedence));
            return true;
        }

        private static bool ApplyReferenceOperator(Stack<FormulaExpression> stack, string op, bool groupResult) {
            if (stack.Count < 2) {
                return false;
            }

            FormulaExpression right = stack.Pop();
            FormulaExpression left = stack.Pop();
            string text = left.Text + op + right.Text;
            stack.Push(new FormulaExpression(groupResult ? "(" + text + ")" : text, 4));
            return true;
        }

        private static bool ApplyUnaryOperator(Stack<FormulaExpression> stack, string op) {
            if (stack.Count < 1) {
                return false;
            }

            FormulaExpression value = stack.Pop();
            string text = value.Precedence < 4 ? "(" + value.Text + ")" : value.Text;
            stack.Push(new FormulaExpression(op + text, 4));
            return true;
        }

        private static bool ApplyPercentOperator(Stack<FormulaExpression> stack) {
            if (stack.Count < 1) {
                return false;
            }

            FormulaExpression value = stack.Pop();
            string text = value.Precedence < 4 ? "(" + value.Text + ")" : value.Text;
            stack.Push(new FormulaExpression(text + "%", 4));
            return true;
        }

        private static bool ApplyParentheses(Stack<FormulaExpression> stack) {
            if (stack.Count < 1) {
                return false;
            }

            FormulaExpression value = stack.Pop();
            stack.Push(new FormulaExpression("(" + value.Text + ")", 4));
            return true;
        }

        private static bool ApplySumAttribute(Stack<FormulaExpression> stack) {
            if (stack.Count < 1) {
                return false;
            }

            FormulaExpression value = stack.Pop();
            stack.Push(new FormulaExpression("SUM(" + value.Text + ")", 4));
            return true;
        }

        private static bool ApplyFixedFunction(Stack<FormulaExpression> stack, ushort functionId) {
            if (!BiffFormulaFunctionMetadata.TryGetFixedFunctionMetadata(functionId, out string? functionName, out int parameterCount) || stack.Count < parameterCount) {
                return false;
            }

            var arguments = new string[parameterCount];
            for (int i = parameterCount - 1; i >= 0; i--) {
                arguments[i] = stack.Pop().Text;
            }

            stack.Push(new FormulaExpression(functionName + "(" + string.Join(",", arguments) + ")", 4));
            return true;
        }

        private static bool ApplyVariableFunction(Stack<FormulaExpression> stack, byte parameterCount, ushort functionBits) {
            bool isCetabFunction = (functionBits & 0x8000) != 0;
            ushort functionId = (ushort)(functionBits & 0x7fff);
            if (!isCetabFunction && functionId == 0x00ff) {
                return ApplyUserDefinedFunction(stack, parameterCount);
            }

            if (isCetabFunction
                || !BiffFormulaFunctionMetadata.TryGetFunctionName(functionId, out string? functionName)
                || !BiffFormulaFunctionMetadata.IsSupportedVariableFunctionArgumentCount(functionId, parameterCount)
                || stack.Count < parameterCount) {
                return false;
            }

            var arguments = new string[parameterCount];
            for (int i = parameterCount - 1; i >= 0; i--) {
                arguments[i] = stack.Pop().Text;
            }

            stack.Push(new FormulaExpression(functionName + "(" + string.Join(",", arguments) + ")", 4));
            return true;
        }

        private static bool ApplyUserDefinedFunction(Stack<FormulaExpression> stack, byte parameterCount) {
            if (parameterCount == 0 || stack.Count < parameterCount) {
                return false;
            }

            var arguments = new string[parameterCount];
            for (int i = parameterCount - 1; i >= 0; i--) {
                arguments[i] = stack.Pop().Text;
            }

            string functionName = arguments[0];
            if (string.IsNullOrWhiteSpace(functionName)) {
                return false;
            }

            stack.Push(new FormulaExpression(functionName + "(" + string.Join(",", arguments.Skip(1)) + ")", 4));
            return true;
        }

        private static string ReadCellReference(byte[] bytes, int offset) {
            ushort row = BiffRecordReader.ReadUInt16(bytes, offset);
            ushort columnBits = BiffRecordReader.ReadUInt16(bytes, offset + 2);
            return BiffFormulaReferenceFormatter.FormatCellReference(row, columnBits);
        }

        private static string ReadAreaReference(byte[] bytes, int offset) {
            ushort firstRow = BiffRecordReader.ReadUInt16(bytes, offset);
            ushort lastRow = BiffRecordReader.ReadUInt16(bytes, offset + 2);
            ushort firstColumnBits = BiffRecordReader.ReadUInt16(bytes, offset + 4);
            ushort lastColumnBits = BiffRecordReader.ReadUInt16(bytes, offset + 6);
            return BiffFormulaReferenceFormatter.FormatAreaReference(firstRow, lastRow, firstColumnBits, lastColumnBits);
        }

        private static string ReadRelativeCellReference(byte[] bytes, int offset, int formulaRow, int formulaColumn) {
            ushort row = BiffRecordReader.ReadUInt16(bytes, offset);
            ushort columnBits = BiffRecordReader.ReadUInt16(bytes, offset + 2);
            int resolvedRow = ResolveRelativeRow(row, columnBits, formulaRow);
            int resolvedColumn = ResolveRelativeColumn(columnBits, formulaColumn);
            return FormatResolvedCellReference(resolvedRow, resolvedColumn, columnBits);
        }

        private static string ReadRelativeAreaReference(byte[] bytes, int offset, int formulaRow, int formulaColumn) {
            ushort firstRow = BiffRecordReader.ReadUInt16(bytes, offset);
            ushort lastRow = BiffRecordReader.ReadUInt16(bytes, offset + 2);
            ushort firstColumnBits = BiffRecordReader.ReadUInt16(bytes, offset + 4);
            ushort lastColumnBits = BiffRecordReader.ReadUInt16(bytes, offset + 6);
            string firstReference = FormatResolvedCellReference(
                ResolveRelativeRow(firstRow, firstColumnBits, formulaRow),
                ResolveRelativeColumn(firstColumnBits, formulaColumn),
                firstColumnBits);
            string lastReference = FormatResolvedCellReference(
                ResolveRelativeRow(lastRow, lastColumnBits, formulaRow),
                ResolveRelativeColumn(lastColumnBits, formulaColumn),
                lastColumnBits);
            return firstReference + ":" + lastReference;
        }

        private static int ResolveRelativeRow(ushort rowValue, ushort columnBits, int formulaRow) {
            bool rowRelative = (columnBits & 0x8000) != 0;
            if (!rowRelative) {
                return rowValue;
            }

            int offset = unchecked((short)rowValue);
            return NormalizeLegacyRow(formulaRow + offset);
        }

        private static int ResolveRelativeColumn(ushort columnBits, int formulaColumn) {
            int columnValue = columnBits & 0x3fff;
            bool columnRelative = (columnBits & 0x4000) != 0;
            if (!columnRelative) {
                return columnValue;
            }

            if ((columnValue & 0x2000) != 0) {
                columnValue -= 0x4000;
            }

            return NormalizeLegacyColumn(formulaColumn + columnValue);
        }

        private static int NormalizeLegacyRow(int row) {
            while (row < 0) {
                row += 65536;
            }

            while (row > 65535) {
                row -= 65536;
            }

            return row;
        }

        private static int NormalizeLegacyColumn(int column) {
            while (column < 0) {
                column += 256;
            }

            while (column > 255) {
                column -= 256;
            }

            return column;
        }

        private static string FormatResolvedCellReference(int zeroBasedRow, int zeroBasedColumn, ushort columnBits) {
            bool rowRelative = (columnBits & 0x8000) != 0;
            bool columnRelative = (columnBits & 0x4000) != 0;
            string columnPrefix = columnRelative ? string.Empty : "$";
            string rowPrefix = rowRelative ? string.Empty : "$";
            return columnPrefix
                + A1.ColumnIndexToLetters(zeroBasedColumn + 1)
                + rowPrefix
                + (zeroBasedRow + 1).ToString(CultureInfo.InvariantCulture);
        }

        private readonly struct FormulaExpression {
            internal FormulaExpression(string text, int precedence) {
                Text = text;
                Precedence = precedence;
            }

            internal string Text { get; }

            internal int Precedence { get; }
        }
    }
}
