using System;

namespace OfficeIMO.Excel.Legacy;

internal enum LaterNumberKind { Double64, Single32, Extended80, Compact16, Compact32 }

internal static class LaterSpreadsheetNumbers {
    internal static double Read(byte[] bytes, int offset, LaterNumberKind kind) {
        if (kind == LaterNumberKind.Compact16) {
            ushort bits = U16(bytes, offset);
            if ((bits & 1) == 0) return (short)bits >> 1;
            int mantissa = bits >> 4; if ((mantissa & 0x800) != 0) mantissa -= 0x1000;
            double[] scales = { 5000, 500, .05, .005, .0005, .00005, 1D / 16, 1D / 64 };
            return mantissa * scales[(bits & 15) >> 1];
        }
        if (kind == LaterNumberKind.Compact32) {
            uint bits = U32(bytes, offset); double mantissa = bits >> 6;
            if ((bits & 32) != 0) mantissa = -mantissa;
            double scale = Math.Pow(10, bits & 15);
            return (bits & 16) != 0 ? mantissa / scale : mantissa * scale;
        }
        if (kind == LaterNumberKind.Extended80) {
            ulong significand = U32(bytes, offset) | ((ulong)U32(bytes, offset + 4) << 32);
            ushort exponent = U16(bytes, offset + 8);
            int magnitude = exponent & 0x7fff;
            if (magnitude == 0x7fff) return double.NaN;
            double value = significand / Math.Pow(2, 63) * Math.Pow(2, (magnitude == 0 ? 1 : magnitude) - 16383);
            return (exponent & 0x8000) == 0 ? value : -value;
        }
        int size = kind == LaterNumberKind.Single32 ? 4 : 8;
        if (BitConverter.IsLittleEndian) return size == 4 ? BitConverter.ToSingle(bytes, offset) : BitConverter.ToDouble(bytes, offset);
        var copy = new byte[size]; Array.Copy(bytes, offset, copy, 0, size); Array.Reverse(copy);
        return size == 4 ? BitConverter.ToSingle(copy, 0) : BitConverter.ToDouble(copy, 0);
    }
    private static ushort U16(byte[] bytes, int p) => (ushort)(bytes[p] | bytes[p + 1] << 8);
    private static uint U32(byte[] bytes, int p) => (uint)(bytes[p] | bytes[p + 1] << 8 | bytes[p + 2] << 16 | bytes[p + 3] << 24);
}
