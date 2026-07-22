using OfficeIMO.Excel.LegacyXls.Write;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Reuses the established formula parser and converts its common BIFF8 token subset to BIFF12.
    /// </summary>
    internal static class XlsbFormulaEncoder {
        private const int MaximumFormulaCharacters = 8_192;
        private const int MaximumControlFunctionDepth = 64;
        private const int MaximumFormulaOperators = 256;

        internal static bool TryEncode(string formulaText, out byte[] formulaPayload, out string? reason) {
            formulaPayload = Array.Empty<byte>();
            if (formulaText == null) throw new ArgumentNullException(nameof(formulaText));
        if (formulaText.Length > MaximumFormulaCharacters) {
                reason = $"Formula text exceeds the supported limit of {MaximumFormulaCharacters} characters.";
            return false;
        }
        if (ExceedsSyntacticNestingLimit(formulaText)) {
            reason = $"Formula syntactic nesting exceeds the supported depth of {MaximumControlFunctionDepth}.";
            return false;
        }
        if (!TryEncodeTokens(formulaText, depth: 0, out byte[] biff12Tokens, out reason)) return false;

            using var payload = new MemoryStream(checked(biff12Tokens.Length + 10));
            WriteUInt16(payload, 0); // grbit
            WriteUInt32(payload, checked((uint)biff12Tokens.Length));
            payload.Write(biff12Tokens, 0, biff12Tokens.Length);
            WriteUInt32(payload, 0); // cbExtra
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

        private static bool TryEncodeTokens(string formulaText, int depth,
            out byte[] tokens, out string? reason) {
            if (depth > MaximumControlFunctionDepth) {
                tokens = Array.Empty<byte>();
                reason = $"Formula control-function nesting exceeds the supported depth of {MaximumControlFunctionDepth}.";
                return false;
            }
            if (TryEncodeControlFunction(formulaText, depth, out tokens, out reason)) return true;
            if (reason != null) return false;

            if (!LegacyXlsFormulaEncoder.TryEncode(formulaText, out byte[] biff8Tokens, out reason)) {
                tokens = Array.Empty<byte>();
                return false;
            }

            return TryConvertTokens(biff8Tokens, out tokens, out reason);
        }

        private static bool TryEncodeControlFunction(string formulaText, int depth,
            out byte[] tokens, out string? reason) {
            tokens = Array.Empty<byte>();
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
            foreach (string argument in arguments) {
                if (argument.Length == 0) {
                    output.WriteByte(0x16); // PtgMissArg
                    continue;
                }

                if (!TryEncodeTokens(argument, checked(depth + 1),
                        out byte[] argumentTokens, out reason)) return false;
                output.Write(argumentTokens, 0, argumentTokens.Length);
            }

            output.WriteByte(0x42); // PtgFuncVarV
            output.WriteByte(checked((byte)arguments.Count));
            WriteUInt16(output, functionId);
            tokens = output.ToArray();
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

        private static bool TryConvertTokens(byte[] source, out byte[] converted, out string? reason) {
            converted = Array.Empty<byte>();
            reason = null;
            using var output = new MemoryStream(source.Length + 16);
            int offset = 0;
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
                    case 0x24:
                    case 0x44:
                    case 0x64:
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
                        reason = $"Formula token 0x{token:X2} is not yet supported by native XLSB generation. Defined names, external references, and array constants remain preservation-only.";
                        return false;
                }
            }

            converted = output.ToArray();
            return true;
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
    }
}
