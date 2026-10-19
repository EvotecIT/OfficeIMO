namespace OfficeIMO.Excel.LegacyXls.Biff;

internal static partial class BiffFormulaTokenScanner {
    /// <summary>Rebinds NAME slots without changing token lengths, relative references, or jump offsets.</summary>
    internal static void RewriteNames(byte[] payload, int start, int length, IReadOnlyDictionary<uint, uint?> names) {
        if (start < 0 || length < 0 || start > payload.Length - length) throw new InvalidDataException("The native formula token range is invalid.");
        int end = start + length, offset = start;
        while (offset < end) {
            int tokenOffset = offset; byte token = payload[offset++];
            int operand = offset;
            bool complete;
            if (IsFixedFunctionToken(token)) complete = TrySkipBytes(ref offset, end, 2);
            else if (IsVariableFunctionToken(token)) complete = TrySkipBytes(ref offset, end, 3);
            else if (token == 0x19) {
                complete = TrySkipBytes(ref offset, end, 1) && TrySkipAttributePayload(payload, ref offset, end, payload[operand]);
            } else complete = TrySkipTokenPayload(payload, ref offset, end, token);
            if (!complete) throw new NotSupportedException("A native formula cannot be rebound without understanding its complete token stream.");
            if (token is not (0x23 or 0x43 or 0x63)) continue;
            uint index = BiffRecordReader.ReadUInt32(payload, operand);
            if (!names.TryGetValue(index, out uint? replacement)) throw new InvalidDataException("A native formula refers to an undeclared defined-name slot.");
            if (replacement.HasValue) {
                uint value = replacement.Value;
                for (int byteIndex = 0; byteIndex < 4; byteIndex++) payload[operand + byteIndex] = (byte)(value >> (byteIndex * 8));
            } else {
                payload[tokenOffset] = (byte)((token & 0x60) | 0x0a); // PtgRefErr has the same operand width as PtgName.
                Array.Clear(payload, operand, 4);
            }
        }
    }
}
