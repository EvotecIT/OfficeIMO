namespace OfficeIMO.Drawing;

/// <summary>Fixed-size MathConstants reader. Every read stays inside the owning MATH table.</summary>
internal static class OfficeOpenTypeMathConstants {
    internal static OfficeMathFontConstants? TryRead(byte[] data, int offset, int length, int unitsPerEm) {
        const int constantsLength = 214; // Four 16-bit values, 51 MathValueRecords, final percentage.
        if (unitsPerEm <= 0 || offset < 0 || length < 10 || offset > data.Length - length) return null;
        if (U16(data, offset) != 1 || U16(data, offset + 2) != 0) return null;
        int relative = U16(data, offset + 4);
        if (relative < 10 || relative > length - constantsLength) return null;
        int start = offset + relative;
        var values = new int[56];
        values[0] = I16(data, start);
        values[1] = I16(data, start + 2);
        values[2] = U16(data, start + 4);
        values[3] = U16(data, start + 6);
        for (int index = 4; index < 55; index++) values[index] = I16(data, start + 8 + (index - 4) * 4);
        values[55] = I16(data, start + 212);
        // Reject scales that could invert or enlarge script layout. Negative design offsets
        // elsewhere, such as radicalKernAfterDegree, are legitimate.
        if (!OfficeMathFontConstants.ValidScriptScales(values)) return null;
        return new OfficeMathFontConstants(unitsPerEm, values);
    }

    private static int U16(byte[] data, int offset) => (data[offset] << 8) | data[offset + 1];
    private static int I16(byte[] data, int offset) => unchecked((short)U16(data, offset));
}
