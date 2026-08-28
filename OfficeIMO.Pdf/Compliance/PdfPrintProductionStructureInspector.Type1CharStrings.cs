namespace OfficeIMO.Pdf;

internal static partial class PdfPrintProductionStructureInspector {
    private const int MaximumType1CharStringDepth = 16;
    private const int MaximumType1CharStringOperations = 100_000;

    private static bool IsValidType1PrivateProgram(byte[] data, int privateBody, int charStringsBody) {
        if (!TryReadType1LenIv(data, privateBody, charStringsBody, out int lenIv) ||
            !TryReadType1Subrs(data, privateBody, charStringsBody, out Dictionary<int, byte[]> subrs) ||
            !TryReadType1CharStrings(data, charStringsBody, out Dictionary<string, byte[]> charStrings) ||
            !charStrings.ContainsKey(".notdef")) return false;

        foreach (byte[] encrypted in charStrings.Values) {
            if (!TryDecodeType1CharString(encrypted, lenIv, out byte[] program) ||
                !TryExecuteType1CharString(program, subrs, lenIv, isSubroutine: false, depth: 0, new List<int>(), out _)) {
                return false;
            }
        }
        return true;
    }

    private static bool TryReadType1LenIv(byte[] data, int start, int end, out int lenIv) {
        lenIv = 4;
        int offset = IndexOfAscii(data, "/lenIV", start);
        if (offset < 0 || offset >= end) return true;
        offset += 6;
        return TryReadType1Integer(data, ref offset, end, out lenIv) && lenIv >= -1 && lenIv <= 255;
    }

    private static bool TryReadType1Subrs(
        byte[] data,
        int start,
        int end,
        out Dictionary<int, byte[]> subrs) {
        subrs = new Dictionary<int, byte[]>();
        int offset = IndexOfAscii(data, "/Subrs", start);
        if (offset < 0 || offset >= end) return true;
        offset += 6;
        if (!TryReadType1Integer(data, ref offset, end, out int declaredCount) || declaredCount < 0 || declaredCount > 65_535 ||
            !TryReadType1Token(data, ref offset, end, out string arrayToken) || arrayToken != "array") return false;

        while (offset < end) {
            SkipType1WhitespaceAndComments(data, ref offset, end);
            if (offset >= end) break;
            if (data[offset] == (byte)'/') {
                offset++;
                while (offset < end && !IsType1Delimiter(data[offset])) offset++;
                continue;
            }
            int entryStart = offset;
            if (!TryReadType1Token(data, ref offset, end, out string token)) return false;
            if (token != "dup") continue;
            if (!TryReadType1Integer(data, ref offset, end, out int index) || index < 0 || index >= declaredCount ||
                !TryReadType1Integer(data, ref offset, end, out int length) || length <= 0 ||
                !TryReadType1Token(data, ref offset, end, out string reader) || reader is not ("RD" or "-|")) {
                offset = entryStart + 1;
                continue;
            }
            if (offset >= end || !IsAsciiWhitespace(data[offset])) return false;
            offset++;
            if (length > end - offset || subrs.ContainsKey(index)) return false;
            var bytes = new byte[length];
            Buffer.BlockCopy(data, offset, bytes, 0, length);
            offset += length;
            if (!TryReadType1Terminator(data, ref offset, end, allowPut: true)) return false;
            subrs.Add(index, bytes);
        }
        return true;
    }

    private static bool TryReadType1CharStrings(
        byte[] data,
        int start,
        out Dictionary<string, byte[]> charStrings) {
        charStrings = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        int offset = start;
        while (offset < data.Length) {
            SkipType1WhitespaceAndComments(data, ref offset, data.Length);
            if (offset >= data.Length) return false;
            if (data[offset] != (byte)'/') {
                if (!TryReadType1Token(data, ref offset, data.Length, out string token)) return false;
                if (token == "end") return charStrings.Count > 0;
                continue;
            }
            offset++;
            int nameStart = offset;
            while (offset < data.Length && !IsType1Delimiter(data[offset])) offset++;
            if (offset == nameStart) return false;
            string name = System.Text.Encoding.ASCII.GetString(data, nameStart, offset - nameStart);
            if (!TryReadType1Integer(data, ref offset, data.Length, out int length) || length <= 0 ||
                !TryReadType1Token(data, ref offset, data.Length, out string reader) || reader is not ("RD" or "-|") ||
                offset >= data.Length || !IsAsciiWhitespace(data[offset])) return false;
            offset++;
            if (length > data.Length - offset || charStrings.ContainsKey(name)) return false;
            var bytes = new byte[length];
            Buffer.BlockCopy(data, offset, bytes, 0, length);
            offset += length;
            if (!TryReadType1Terminator(data, ref offset, data.Length, allowPut: false)) return false;
            charStrings.Add(name, bytes);
        }
        return false;
    }

    private static bool TryReadType1Terminator(byte[] data, ref int offset, int end, bool allowPut) {
        if (!TryReadType1Token(data, ref offset, end, out string token)) return false;
        if (token is "ND" or "|-") return true;
        if (allowPut && token is "NP" or "put") return true;
        if (token == "def") return true;
        if (token != "noaccess") return false;
        return TryReadType1Token(data, ref offset, end, out token) && token is "def" or "put";
    }

    private static bool TryDecodeType1CharString(byte[] encrypted, int lenIv, out byte[] program) {
        program = Array.Empty<byte>();
        if (lenIv == -1) {
            program = (byte[])encrypted.Clone();
            return program.Length > 0;
        }
        if (encrypted.Length <= lenIv) return false;
        var decrypted = new byte[encrypted.Length];
        ushort state = 4330;
        for (int index = 0; index < encrypted.Length; index++) {
            byte cipher = encrypted[index];
            decrypted[index] = (byte)(cipher ^ (state >> 8));
            state = unchecked((ushort)((cipher + state) * 52845 + 22719));
        }
        program = new byte[decrypted.Length - lenIv];
        Buffer.BlockCopy(decrypted, lenIv, program, 0, program.Length);
        return program.Length > 0;
    }

    private static bool TryExecuteType1CharString(
        byte[] program,
        Dictionary<int, byte[]> encryptedSubrs,
        int lenIv,
        bool isSubroutine,
        int depth,
        List<int> stack,
        out bool terminated) {
        terminated = false;
        if (depth > MaximumType1CharStringDepth) return false;
        int operations = 0;
        for (int offset = 0; offset < program.Length;) {
            if (++operations > MaximumType1CharStringOperations) return false;
            byte value = program[offset++];
            if (value >= 32) {
                if (!TryReadType1CharStringNumber(program, ref offset, value, out int number) || stack.Count >= 48) return false;
                stack.Add(number);
                continue;
            }

            switch (value) {
                case 1:
                case 3:
                    if (!ConsumeType1Arguments(stack, 2, requireMultiple: true)) return false;
                    break;
                case 4:
                case 22:
                    if (!ConsumeType1Arguments(stack, 1)) return false;
                    break;
                case 5:
                    if (!ConsumeType1Arguments(stack, 2, requireMultiple: true)) return false;
                    break;
                case 6:
                case 7:
                    if (!ConsumeType1Arguments(stack, 1, requireAtLeast: true)) return false;
                    break;
                case 8:
                    if (!ConsumeType1Arguments(stack, 6, requireMultiple: true)) return false;
                    break;
                case 9:
                    if (stack.Count != 0) return false;
                    break;
                case 10:
                    if (stack.Count < 1) return false;
                    int subrIndex = stack[stack.Count - 1];
                    stack.RemoveAt(stack.Count - 1);
                    if (!encryptedSubrs.TryGetValue(subrIndex, out byte[]? encryptedSubr) ||
                        !TryDecodeType1CharString(encryptedSubr, lenIv, out byte[] subr) ||
                        !TryExecuteType1CharString(subr, encryptedSubrs, lenIv, true, depth + 1, stack, out bool returned) ||
                        !returned) return false;
                    break;
                case 11:
                    if (!isSubroutine) return false;
                    terminated = true;
                    return true;
                case 12:
                    if (offset >= program.Length || !ExecuteType1Escape(program[offset++], stack, isSubroutine, out bool escapedEnd)) return false;
                    if (escapedEnd) {
                        terminated = true;
                        return offset == program.Length;
                    }
                    break;
                case 13:
                    if (!ConsumeType1Arguments(stack, 2)) return false;
                    break;
                case 14:
                    if (isSubroutine || stack.Count != 0) return false;
                    terminated = true;
                    return offset == program.Length;
                case 21:
                    if (!ConsumeType1Arguments(stack, 2)) return false;
                    break;
                case 30:
                case 31:
                    if (!ConsumeType1Arguments(stack, 4)) return false;
                    break;
                default:
                    return false;
            }
        }
        return false;
    }

    private static bool ExecuteType1Escape(byte operation, List<int> stack, bool isSubroutine, out bool terminated) {
        terminated = false;
        switch (operation) {
            case 0:
                return ConsumeType1Arguments(stack, 0);
            case 1:
            case 2:
                return ConsumeType1Arguments(stack, 6);
            case 6:
                if (isSubroutine || !ConsumeType1Arguments(stack, 5)) return false;
                terminated = true;
                return true;
            case 7:
                return ConsumeType1Arguments(stack, 4);
            case 12:
                if (stack.Count < 2) return false;
                int divisor = stack[stack.Count - 1];
                int dividend = stack[stack.Count - 2];
                stack.RemoveRange(stack.Count - 2, 2);
                stack.Add(divisor == 0 ? 0 : dividend / divisor);
                return divisor != 0;
            case 16:
                if (stack.Count < 2) return false;
                int argumentCount = stack[stack.Count - 2];
                if (argumentCount < 0 || argumentCount > stack.Count - 2) return false;
                stack.RemoveRange(stack.Count - argumentCount - 2, argumentCount + 2);
                return true;
            case 17:
                if (stack.Count >= 48) return false;
                stack.Add(0);
                return true;
            case 33:
                return ConsumeType1Arguments(stack, 2);
            default:
                return false;
        }
    }

    private static bool ConsumeType1Arguments(
        List<int> stack,
        int count,
        bool requireMultiple = false,
        bool requireAtLeast = false) {
        if (requireAtLeast) {
            if (stack.Count < count) return false;
        } else if (requireMultiple) {
            if (stack.Count < count || stack.Count % count != 0) return false;
        } else if (stack.Count != count) {
            return false;
        }
        stack.Clear();
        return true;
    }

    private static bool TryReadType1CharStringNumber(byte[] data, ref int offset, byte first, out int value) {
        value = 0;
        if (first <= 246) {
            value = first - 139;
            return true;
        }
        if (first <= 250) {
            if (offset >= data.Length) return false;
            value = (first - 247) * 256 + data[offset++] + 108;
            return true;
        }
        if (first <= 254) {
            if (offset >= data.Length) return false;
            value = -(first - 251) * 256 - data[offset++] - 108;
            return true;
        }
        if (offset > data.Length - 4) return false;
        value = (data[offset] << 24) | (data[offset + 1] << 16) | (data[offset + 2] << 8) | data[offset + 3];
        offset += 4;
        return true;
    }

    private static bool TryReadType1Integer(byte[] data, ref int offset, int end, out int value) {
        value = 0;
        if (!TryReadType1Token(data, ref offset, end, out string token)) return false;
        return int.TryParse(token, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out value);
    }

    private static bool TryReadType1Token(byte[] data, ref int offset, int end, out string token) {
        token = string.Empty;
        SkipType1WhitespaceAndComments(data, ref offset, end);
        if (offset >= end) return false;
        int start = offset;
        while (offset < end && !IsType1Delimiter(data[offset])) offset++;
        if (offset == start) {
            offset++;
            return false;
        }
        token = System.Text.Encoding.ASCII.GetString(data, start, offset - start);
        return true;
    }

    private static void SkipType1WhitespaceAndComments(byte[] data, ref int offset, int end) {
        while (offset < end) {
            if (IsAsciiWhitespace(data[offset])) {
                offset++;
                continue;
            }
            if (data[offset] != (byte)'%') return;
            while (offset < end && data[offset] is not 10 and not 13) offset++;
        }
    }

    private static bool IsType1Delimiter(byte value) =>
        IsAsciiWhitespace(value) || value is (byte)'(' or (byte)')' or (byte)'<' or (byte)'>' or
            (byte)'[' or (byte)']' or (byte)'{' or (byte)'}' or (byte)'/' or (byte)'%';
}
