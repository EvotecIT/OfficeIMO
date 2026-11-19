using OfficeIMO.Excel.LegacyXls.Write;
using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Reuses the established formula parser and converts its common BIFF8 token subset to BIFF12.
    /// </summary>
    internal static partial class XlsbFormulaEncoder {
        private const int MaximumFormulaCharacters = 8_192;
        private const int MaximumControlFunctionDepth = 64;
        private const int MaximumFormulaOperators = 256;
        private const int MaximumTokenBytes = 16_384;

        internal static bool TryEncode(string formulaText, out byte[] formulaPayload, out string? reason, bool arrayFormula = false) =>
            TryEncode(formulaText, LegacyXlsFormulaNameIndex.Empty, -1, null, out formulaPayload, out reason, arrayFormula);

        /// <remarks>futureFunctions returns the one-based BrtName index of a future-function name (its canonical
        /// _xlfn. spelling), or null when the target workbook cannot provide that name.</remarks>
        internal static bool TryEncode(string formulaText, LegacyXlsFormulaNameIndex names, int sheetIndex, Func<ushort, ushort>? mapExternal,
            out byte[] formulaPayload, out string? reason, bool arrayFormula = false, byte resultClass = 0,
            Func<string, uint?>? futureFunctions = null, XlsbTableReferenceIndex? tables = null) {
            formulaPayload = Array.Empty<byte>();
            ArgumentNullException.ThrowIfNull(formulaText);
            if (!XlsbStructuredReferences.TrySubstitute(formulaText, tables, names, sheetIndex, mapExternal, out formulaText,
                    out IReadOnlyList<byte[]> lists, out reason)) return false;
            var placeholders = lists.Count == 0 ? null : new XlsbStructuredReferences.Placeholders(lists);
            if (formulaText.Length > MaximumFormulaCharacters) {
                reason = $"Formula text exceeds the supported limit of {MaximumFormulaCharacters} characters.";
                return false;
            }
            if (ExceedsSyntacticNestingLimit(formulaText)) {
                reason = $"Formula syntactic nesting exceeds the supported depth of {MaximumControlFunctionDepth}.";
                return false;
            }
            var scope = new EncodeScope(names, sheetIndex, mapExternal, futureFunctions, placeholders);
            if (!TryEncodeTokens(formulaText, depth: 0, scope, out byte[] biff12Tokens, out byte[] extra, out reason)) return false;
            if (placeholders != null && !placeholders.IsComplete) {
                reason = "A structured table reference is used where native XLSB formulas cannot store it.";
                return false;
            }
            if (biff12Tokens.Length == 0 || biff12Tokens.Length > MaximumTokenBytes) {
                reason = "The encoded formula is outside the XLSB formula-token size limit.";
                return false;
            }

            if (resultClass == 0) resultClass = arrayFormula ? ExcelNativeFormulaOperandTypes.Array : ExcelNativeFormulaOperandTypes.Value;
            if (!ExcelNativeFormulaOperandTypes.TryApply(biff12Tokens, binary12: true, resultClass, out reason)) return false;
            using var payload = new MemoryStream(checked(biff12Tokens.Length + extra.Length + 10));
            WriteUInt16(payload, 0); // grbit
            WriteUInt32(payload, checked((uint)biff12Tokens.Length));
            payload.Write(biff12Tokens, 0, biff12Tokens.Length);
            WriteUInt32(payload, checked((uint)extra.Length)); // cb
            payload.Write(extra, 0, extra.Length); // rgcb: one PtgExtraArray per PtgArray, in token order
            formulaPayload = payload.ToArray();
            return true;
        }

        private static bool ExceedsSyntacticNestingLimit(string formulaText) {
            int parenthesisDepth = 0;
            int unaryRun = 0;
            int operators = 0;
            bool quoted = false;
            for (int index = 0; index < formulaText.Length; index++) {
                char current = formulaText[index];
                if (current == '"') {
                    if (quoted && index + 1 < formulaText.Length && formulaText[index + 1] == '"') {
                        index++;
                        continue;
                    }
                    quoted = !quoted;
                    unaryRun = 0;
                    continue;
                }
                if (quoted) continue;
                if (current == '(') {
                    parenthesisDepth++;
                    if (parenthesisDepth > MaximumControlFunctionDepth) return true;
                    unaryRun = 0;
                } else if (current == ')') {
                    if (parenthesisDepth > 0) parenthesisDepth--;
                    unaryRun = 0;
                } else if (current == '+' || current == '-') {
                    unaryRun++;
                    if (unaryRun > MaximumControlFunctionDepth) return true;
                    if (++operators > MaximumFormulaOperators) return true;
                } else if (current == '*' || current == '/' || current == '^' ||
                           current == '&' || current == '=' || current == '<' ||
                           current == '>' || current == '%') {
                    unaryRun = 0;
                    if (++operators > MaximumFormulaOperators) return true;
                } else if (!char.IsWhiteSpace(current)) {
                    unaryRun = 0;
                }
            }
            return false;
        }

        private static bool TryEncodeTokens(string formulaText, int depth, EncodeScope scope,
            out byte[] tokens, out byte[] extra, out string? reason) {
            extra = Array.Empty<byte>();
            if (depth > MaximumControlFunctionDepth) {
                tokens = Array.Empty<byte>();
                reason = $"Formula control-function nesting exceeds the supported depth of {MaximumControlFunctionDepth}.";
                return false;
            }
            if (TryEncodeControlFunction(formulaText, depth, scope, out tokens, out extra, out reason)) return true;
            if (reason != null) return false;
            if (ContainsFutureFunction(formulaText)) {
                return TryEncodeFutureFunctionExpression(formulaText, depth, scope, out tokens, out extra, out reason);
            }

            byte[] biff8Tokens;
            byte[] biff8Extra = Array.Empty<byte>();
            bool encoded = ContainsArrayConstant(formulaText)
                ? LegacyXlsFormulaEncoder.TryEncodeFormulaRecord(formulaText, scope.Names, scope.SheetIndex, out biff8Tokens, out biff8Extra, out reason)
                : LegacyXlsFormulaEncoder.TryEncode(formulaText, scope.Names, scope.SheetIndex, out biff8Tokens, out reason);
            if (!encoded) {
                tokens = Array.Empty<byte>();
                return false;
            }

            if (!TryConvertTokens(biff8Tokens, biff8Extra, scope.MapExternal, out tokens, out extra, out reason)) return false;
            // Swap placeholders in each leaf so control-function jumps are computed over the final token sizes.
            if (scope.Placeholders != null && !scope.Placeholders.TryReplace(tokens, out tokens)) {
                reason = "A structured table reference could not be placed in the native formula.";
                return false;
            }
            return true;
        }

        private static bool TryEncodeControlFunction(string formulaText, int depth, EncodeScope scope,
            out byte[] tokens, out byte[] extra, out string? reason) {
            tokens = Array.Empty<byte>();
            extra = Array.Empty<byte>();
            reason = null;
            string normalized = formulaText.Trim();
            if (normalized.StartsWith("=", StringComparison.Ordinal)) normalized = normalized.Substring(1).Trim();

            int open = normalized.IndexOf('(');
            if (open <= 0 || !normalized.EndsWith(")", StringComparison.Ordinal)) return false;
            string name = normalized.Substring(0, open).Trim();
            ushort functionId;
            int minimumArguments;
            int maximumArguments;
            if (name.Equals("IF", StringComparison.OrdinalIgnoreCase)) {
                functionId = 0x0001;
                minimumArguments = 2;
                maximumArguments = 3;
            } else if (name.Equals("CHOOSE", StringComparison.OrdinalIgnoreCase)) {
                functionId = 0x0064;
                minimumArguments = 2;
                maximumArguments = 30;
            } else if (name.Equals("IFERROR", StringComparison.OrdinalIgnoreCase)) {
                functionId = 0x01E0;
                minimumArguments = 2;
                maximumArguments = 2;
            } else {
                return false;
            }

            string argumentText = normalized.Substring(open + 1, normalized.Length - open - 2);
            if (!TrySplitArguments(argumentText, out IReadOnlyList<string>? arguments)
                || arguments!.Count < minimumArguments
                || arguments.Count > maximumArguments) {
                reason = $"Formula function {name.ToUpperInvariant()} requires between {minimumArguments} and {maximumArguments} arguments.";
                return false;
            }

            using var output = new MemoryStream();
            using var extraOutput = new MemoryStream();
            if (!TryEncodeArguments(arguments, depth, scope, output, extraOutput, out reason)) return false;
            output.WriteByte(0x42); // PtgFuncVar
            output.WriteByte(checked((byte)arguments.Count));
            WriteUInt16(output, functionId);
            tokens = output.ToArray();
            extra = extraOutput.ToArray();
            return true;
        }

        private static bool TryEncodeArguments(IReadOnlyList<string> arguments, int depth, EncodeScope scope,
            Stream output, Stream extraOutput, out string? reason) {
            reason = null;
            foreach (string argument in arguments) {
                if (argument.Length == 0) {
                    output.WriteByte(0x16); // PtgMissArg
                    continue;
                }

                if (!TryEncodeTokens(argument, checked(depth + 1), scope,
                        out byte[] argumentTokens, out byte[] argumentExtra, out reason)) return false;
                output.Write(argumentTokens, 0, argumentTokens.Length);
                extraOutput.Write(argumentExtra, 0, argumentExtra.Length);
            }
            return true;
        }

        private static bool TrySplitArguments(string text, out IReadOnlyList<string>? arguments) {
            arguments = null;
            var result = new List<string>();
            var current = new StringBuilder(text.Length);
            bool inString = false;
            bool inQuotedSheet = false;
            int parenthesisDepth = 0;
            int arrayDepth = 0;
            for (int index = 0; index < text.Length; index++) {
                char character = text[index];
                if (inString) {
                    current.Append(character);
                    if (character == '"') {
                        if (index + 1 < text.Length && text[index + 1] == '"') {
                            current.Append(text[++index]);
                        } else {
                            inString = false;
                        }
                    }
                    continue;
                }

                if (inQuotedSheet) {
                    current.Append(character);
                    if (character == '\'' && index + 1 < text.Length && text[index + 1] == '\'') {
                        current.Append(text[++index]);
                    } else if (character == '\'') {
                        inQuotedSheet = false;
                    }
                    continue;
                }

                if (character == '"') {
                    inString = true;
                } else if (character == '\'') {
                    inQuotedSheet = true;
                } else if (character == '(') {
                    parenthesisDepth++;
                } else if (character == ')') {
                    if (--parenthesisDepth < 0) return false;
                } else if (character == '{') {
                    arrayDepth++;
                } else if (character == '}') {
                    if (--arrayDepth < 0) return false;
                } else if (character == ',' && parenthesisDepth == 0 && arrayDepth == 0) {
                    result.Add(current.ToString().Trim());
                    current.Clear();
                    continue;
                }
                current.Append(character);
            }

            if (inString || inQuotedSheet || parenthesisDepth != 0 || arrayDepth != 0) return false;
            result.Add(current.ToString().Trim());
            arguments = result;
            return true;
        }

        private static bool TryConvertTokens(byte[] source, byte[] sourceExtra, Func<ushort, ushort>? mapExternal,
            out byte[] converted, out byte[] convertedExtra, out string? reason) {
            converted = Array.Empty<byte>();
            convertedExtra = Array.Empty<byte>();
            reason = null;
            using var output = new MemoryStream(source.Length + 16);
            using var extraOutput = new MemoryStream(sourceExtra.Length + 16);
            int offset = 0;
            int extraOffset = 0;
            while (offset < source.Length) {
                byte token = source[offset++];
                switch (token) {
                    case >= 0x03 and <= 0x16:
                        output.WriteByte(token);
                        break;
                    case 0x17:
                        if (!TryConvertString(source, ref offset, output)) {
                            reason = "Formula contains an invalid BIFF string token.";
                            return false;
                        }
                        break;
                    case 0x19:
                        if (!TryConvertAttribute(source, ref offset, output, out reason)) return false;
                        break;
                    case 0x1C:
                    case 0x1D:
                        if (!TryCopy(source, ref offset, output, token, 1)) return InvalidToken(out reason);
                        break;
                    case 0x1E:
                        if (!TryCopy(source, ref offset, output, token, 2)) return InvalidToken(out reason);
                        break;
                    case 0x1F:
                        if (!TryCopy(source, ref offset, output, token, 8)) return InvalidToken(out reason);
                        break;
                    case 0x20:
                    case 0x40:
                    case 0x60:
                        // BIFF8 PtgArray has 7 unused bytes; BIFF12 PtgArray (2.5.98.23) has 14.
                        if (offset > source.Length - 7) return InvalidToken(out reason);
                        offset += 7;
                        output.WriteByte(token);
                        output.Write(new byte[14], 0, 14);
                        if (!TryConvertArrayConstant(sourceExtra, ref extraOffset, extraOutput, out reason)) return false;
                        break;
                    case 0x21:
                    case 0x41:
                    case 0x61:
                        if (!TryCopy(source, ref offset, output, token, 2)) return InvalidToken(out reason);
                        break;
                    case 0x22:
                    case 0x42:
                    case 0x62:
                        if (!TryCopy(source, ref offset, output, token, 3)) return InvalidToken(out reason);
                        break;
                    case 0x23:
                    case 0x43:
                    case 0x63:
                        if (!TryCopy(source, ref offset, output, token, 4)) return InvalidToken(out reason);
                        break;
                    case 0x3A:
                    case 0x5A:
                    case 0x7A:
                    case 0x3B:
                    case 0x5B:
                    case 0x7B:
                        if (offset > source.Length - 2) return InvalidToken(out reason);
                        ushort external = ReadUInt16(source, offset); offset += 2;
                        using (var reference = new MemoryStream()) {
                            bool area = (token & 0x1F) == 0x1B;
                            if (!(area ? TryConvertArea(source, ref offset, reference, token) : TryConvertReference(source, ref offset, reference, token))) return InvalidToken(out reason);
                            byte[] bytes = reference.ToArray(); output.WriteByte(bytes[0]);
                            WriteUInt16(output, mapExternal?.Invoke(external) ?? external);
                            output.Write(bytes, 1, bytes.Length - 1);
                        }
                        break;
                    case 0x24:
                    case 0x44:
                    case 0x64:
                        if (TryConvertCellRange(source, ref offset, output)) break;
                        if (!TryConvertReference(source, ref offset, output, token)) return InvalidToken(out reason);
                        break;
                    case 0x2C:
                    case 0x4C:
                    case 0x6C:
                        if (!TryConvertReference(source, ref offset, output, token)) return InvalidToken(out reason);
                        break;
                    case 0x25:
                    case 0x45:
                    case 0x65:
                    case 0x2D:
                    case 0x4D:
                    case 0x6D:
                        if (!TryConvertArea(source, ref offset, output, token)) return InvalidToken(out reason);
                        break;
                    default:
                        reason = $"Formula token 0x{token:X2} is not yet supported by native XLSB generation. Defined names and external references remain preservation-only.";
                        return false;
                }
            }
            if (extraOffset != sourceExtra.Length) {
                reason = "Formula array-constant data does not match its array tokens.";
                return false;
            }

            converted = output.ToArray();
            convertedExtra = extraOutput.ToArray();
            return true;
        }

        /// <summary>Converts one BIFF8 array-constant block to a BIFF12 PtgExtraArray (2.5.98.41) of SerAr values.</summary>
        private static bool TryConvertArrayConstant(byte[] source, ref int offset, Stream output, out string? reason) {
            reason = "Formula array-constant data is truncated or invalid.";
            if (offset > source.Length - 3) return false;
            int columns = source[offset] + 1;
            int rows = (source[offset + 1] | source[offset + 2] << 8) + 1;
            offset += 3;
            WriteUInt32(output, checked((uint)rows));
            WriteUInt32(output, checked((uint)columns));
            for (int index = 0; index < rows * columns; index++) {
                if (offset >= source.Length) return false;
                byte kind = source[offset++];
                switch (kind) {
                    case 0x01: // number -> SerNum
                        if (offset > source.Length - 8) return false;
                        output.WriteByte(0x00);
                        output.Write(source, offset, 8);
                        offset += 8;
                        break;
                    case 0x02: { // string -> SerStr
                        if (offset > source.Length - 3) return false;
                        int length = source[offset] | source[offset + 1] << 8;
                        bool wide = (source[offset + 2] & 0x01) != 0;
                        offset += 3;
                        int byteCount = checked(length * (wide ? 2 : 1));
                        if (offset > source.Length - byteCount) return false;
                        string value = wide ? Encoding.Unicode.GetString(source, offset, byteCount) : Latin1(source, offset, byteCount);
                        offset += byteCount;
                        if (value.Length > 255) {
                            reason = "Formula array constants support strings of up to 255 characters in native XLSB generation.";
                            return false;
                        }
                        output.WriteByte(0x01);
                        WriteUInt16(output, checked((ushort)value.Length));
                        byte[] bytes = Encoding.Unicode.GetBytes(value);
                        output.Write(bytes, 0, bytes.Length);
                        break;
                    }
                    case 0x04: // Boolean -> SerBool
                        if (offset > source.Length - 8) return false;
                        output.WriteByte(0x02);
                        output.WriteByte(source[offset] == 0 ? (byte)0 : (byte)1);
                        offset += 8;
                        break;
                    case 0x10: // error -> SerErr
                        if (offset > source.Length - 8) return false;
                        output.WriteByte(0x04);
                        output.WriteByte(source[offset]);
                        output.WriteByte(0);
                        WriteUInt16(output, 0);
                        offset += 8;
                        break;
                    case 0x00:
                        reason = "Native XLSB formulas cannot store empty array-constant elements.";
                        return false;
                    default:
                        return false;
                }
            }
            reason = null;
            return true;
        }

        private static string Latin1(byte[] source, int offset, int count) {
            var characters = new char[count];
            for (int index = 0; index < count; index++) characters[index] = (char)source[offset + index];
            return new string(characters);
        }

        private static bool ContainsArrayConstant(string formulaText) {
            bool quoted = false;
            foreach (char character in formulaText) {
                if (character == '"') quoted = !quoted;
                else if (character == '{' && !quoted) return true;
            }
            return false;
        }

        private static bool TryConvertString(byte[] source, ref int offset, Stream output) {
            if (offset > source.Length - 2) return false;
            int characterCount = source[offset++];
            byte flags = source[offset++];
            bool wide = (flags & 0x01) != 0;
            if ((flags & 0xFE) != 0) return false;
            int byteCount = checked(characterCount * (wide ? 2 : 1));
            if (offset > source.Length - byteCount) return false;

            string value = wide
                ? Encoding.Unicode.GetString(source, offset, byteCount)
                : Encoding.ASCII.GetString(source, offset, byteCount);
            offset += byteCount;
            output.WriteByte(0x17);
            WriteUInt16(output, checked((ushort)value.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
            return true;
        }

        private static bool TryConvertAttribute(
            byte[] source,
            ref int offset,
            Stream output,
            out string? reason) {
            reason = null;
            if (offset > source.Length - 3) return InvalidToken(out reason);
            byte attribute = source[offset++];
            if (attribute == 0x02 || attribute == 0x04 || attribute == 0x08 || attribute == 0x80) {
                reason = "IF, CHOOSE, and other jump-optimized formula token streams are not yet supported by native XLSB generation.";
                return false;
            }

            output.WriteByte(0x19);
            output.WriteByte(attribute);
            output.WriteByte(source[offset++]);
            output.WriteByte(source[offset++]);
            return true;
        }

        private static bool TryConvertReference(byte[] source, ref int offset, Stream output, byte token) {
            if (offset > source.Length - 4) return false;
            ushort row = ReadUInt16(source, offset);
            ushort columnBits = ReadUInt16(source, offset + 2);
            offset += 4;
            output.WriteByte(token);
            WriteUInt32(output, row);
            WriteUInt16(output, columnBits);
            return true;
        }

        private static bool TryConvertArea(byte[] source, ref int offset, Stream output, byte token) {
            if (offset > source.Length - 8) return false;
            ushort firstRow = ReadUInt16(source, offset);
            ushort lastRow = ReadUInt16(source, offset + 2);
            ushort firstColumnBits = ReadUInt16(source, offset + 4);
            ushort lastColumnBits = ReadUInt16(source, offset + 6);
            offset += 8;
            output.WriteByte(NormalizeAreaToken(token));
            WriteUInt32(output, firstRow);
            WriteUInt32(output, lastRow);
            WriteUInt16(output, firstColumnBits);
            WriteUInt16(output, lastColumnBits);
            return true;
        }

        /// <summary>
        /// Writes "PtgRef PtgRef PtgRange" for a plain A1:B2 range as the PtgArea Excel stores for it. Excel evaluates a
        /// range operator result in value context as #VALUE!, where the same area is implicitly intersected.
        /// </summary>
        private static bool TryConvertCellRange(byte[] source, ref int offset, Stream output) {
            int second = offset + 4;
            if (second > source.Length - 6 || source[second] is not (0x24 or 0x44 or 0x64) || source[second + 5] != 0x11) return false;
            ushort firstRow = ReadUInt16(source, offset);
            ushort firstColumnBits = ReadUInt16(source, offset + 2);
            ushort lastRow = ReadUInt16(source, second + 1);
            ushort lastColumnBits = ReadUInt16(source, second + 3);
            if (firstRow > lastRow || (firstColumnBits & 0x3FFF) > (lastColumnBits & 0x3FFF)) return false;
            output.WriteByte(0x25); // PtgArea; the operand class is assigned after conversion
            WriteUInt32(output, firstRow);
            WriteUInt32(output, lastRow);
            WriteUInt16(output, firstColumnBits);
            WriteUInt16(output, lastColumnBits);
            offset = second + 6;
            return true;
        }

        private static bool TryCopy(byte[] source, ref int offset, Stream output, byte token, int count) {
            if (offset > source.Length - count) return false;
            output.WriteByte(token);
            output.Write(source, offset, count);
            offset += count;
            return true;
        }

        private static byte NormalizeAreaToken(byte token) => token switch {
            0x45 or 0x65 => 0x25,
            0x4D or 0x6D => 0x2D,
            _ => token,
        };

        private static bool InvalidToken(out string? reason) {
            reason = "Formula contains a truncated or invalid BIFF token stream.";
            return false;
        }

        private static ushort ReadUInt16(byte[] source, int offset) =>
            checked((ushort)(source[offset] | (source[offset + 1] << 8)));

        private static void WriteUInt16(Stream stream, ushort value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
        }

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)(value >> 16));
            stream.WriteByte((byte)(value >> 24));
        }

        private sealed class EncodeScope {
            internal EncodeScope(LegacyXlsFormulaNameIndex names, int sheetIndex, Func<ushort, ushort>? mapExternal, Func<string, uint?>? futureFunctions,
                XlsbStructuredReferences.Placeholders? placeholders) {
                Names = names;
                SheetIndex = sheetIndex;
                MapExternal = mapExternal;
                FutureFunctions = futureFunctions;
                Placeholders = placeholders;
            }

            internal LegacyXlsFormulaNameIndex Names { get; }
            internal int SheetIndex { get; }
            internal Func<ushort, ushort>? MapExternal { get; }
            internal Func<string, uint?>? FutureFunctions { get; }
            internal XlsbStructuredReferences.Placeholders? Placeholders { get; }
        }
    }
}
