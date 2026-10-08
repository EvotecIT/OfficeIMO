using OfficeIMO.Drawing;
using static OfficeIMO.Access.AccessNativeBinary;

namespace OfficeIMO.Access;

internal sealed partial class AccessNativeDatabase {
    internal object DecodeScalar(AccessNativeColumn column, OfficeByteView bytes, CancellationToken cancellation, int? maximumBytes = null) {
        cancellation.ThrowIfCancellationRequested();
        int limit = maximumBytes ?? MaxValueBytes;
        if (bytes.Length > limit) throw new InvalidDataException("Native Access value exceeds its value limit.");
        if (maximumBytes.HasValue && column.Type != 11 && column.Type != 12) AccountMetadata(bytes.Length);
        if (column.Calculated) return new AccessOpaqueValue(column.Type, bytes.ToArray(), "Calculated native field representation is retained; expressions are never evaluated.");
        int expected = column.Type switch { 1 => 1, 2 => 1, 3 => 2, 4 => 4, 5 => 8, 6 => 4, 7 => 8, 8 => 8, 15 => 16, 19 => 8, 16 => 17, 18 => 4, 20 => 42, _ => -1 };
        if (expected >= 0 && bytes.Length != expected) throw new InvalidDataException("Native Access fixed value has an invalid byte length.");
        switch (column.Type) {
            case 1: return bytes[0] != 0;
            case 2: return bytes[0];
            case 3: return checked((short)I16(bytes, 0));
            case 4: return I32(bytes, 0);
            case 5: return I64(bytes, 0) / 10000m;
            case 6: return BitConverter.ToSingle(bytes.ToArray(), 0);
            case 7: return F64(bytes, 0);
            case 8:
                try { return DateTime.SpecifyKind(DateTime.FromOADate(F64(bytes, 0)), DateTimeKind.Unspecified); }
                catch (ArgumentException exception) { throw new InvalidDataException("Native Access date is outside its supported range.", exception); }
            case 9: case 17: return bytes.ToArray();
            case 10: { string value = Text(bytes); return column.RedactConnection ? RedactConnection(value)! : value; }
            case 11: return LongValue(bytes, cancellation, maximumBytes);
            case 12: { string value = Text(LongValue(bytes, cancellation, maximumBytes)); return column.RedactConnection ? RedactConnection(value)! : value; }
            case 15: return new Guid(bytes.ToArray());
            case 19: return I64(bytes, 0);
            case 16: return Numeric(column, bytes);
            case 20: return ExtendedDate(bytes);
            case 18: return column.ComplexDefinition == null ? new AccessOpaqueValue(column.Type, bytes.ToArray(), "Unqualified complex definition retains its exact foreign-key bytes.") : new AccessComplexValue(column.ComplexDefinition, I32(bytes, 0));
            default: return new AccessOpaqueValue(column.Type, bytes.ToArray(), "This native field type has no qualified typed value decoder.");
        }
    }
    private static DateTime ExtendedDate(OfficeByteView bytes) {
        if (bytes[19] != ':' || bytes[39] != ':' || bytes[40] < '0' || bytes[40] > '7' || bytes[41] != 0) throw new InvalidDataException("Native Access extended-date separators are invalid.");
        long days = Digits(bytes, 0, 19), units = Digits(bytes, 20, 19); int scale = bytes[40] - '0';
        try {
            long multiplier = 1; for (int i = scale; i < 7; i++) multiplier *= 10;
            long ticks = checked(checked(days * TimeSpan.TicksPerDay) + checked(units * multiplier));
            return new DateTime(ticks, DateTimeKind.Unspecified);
        } catch (Exception exception) when (exception is OverflowException || exception is ArgumentOutOfRangeException) { throw new InvalidDataException("Native Access extended date exceeds the CLR DateTime range.", exception); }
    }
    private static long Digits(OfficeByteView bytes, int start, int count) {
        long value = 0;
        try { for (int i = start; i < start + count; i++) { if (bytes[i] < '0' || bytes[i] > '9') throw new InvalidDataException("Native Access extended date contains a non-digit."); value = checked(value * 10 + bytes[i] - '0'); } }
        catch (OverflowException exception) { throw new InvalidDataException("Native Access extended date magnitude is invalid.", exception); }
        return value;
    }
    private static object Numeric(AccessNativeColumn column, OfficeByteView bytes) {
        if (column.Precision < 1 || column.Scale > 28 || column.Precision > 28) return new AccessOpaqueValue(column.Type, bytes.ToArray(), "The numeric definition has no qualified CLR Decimal precision; its exact native representation is retained.");
        try {
            decimal magnitude = 0;
            for (int position = 1; position < 17; position += 4) magnitude = checked(magnitude * 4294967296m + U32(bytes, position));
            for (int i = 0; i < column.Scale; i++) magnitude /= 10m;
            return bytes[0] == 0 ? magnitude : -magnitude;
        } catch (OverflowException exception) { throw new InvalidDataException("Native Access numeric magnitude exceeds its qualified Decimal definition.", exception); }
    }
    internal byte[] LongValue(OfficeByteView definition, CancellationToken cancellation, int? maximumBytes = null) {
        using var stream = new AccessNativeLongValueStream(this, definition, cancellation, maximumBytes);
        if (maximumBytes.HasValue) AccountMetadata(checked((int)stream.Length));
        var output = new byte[checked((int)stream.Length)]; int copied = 0;
        while (copied < output.Length) { cancellation.ThrowIfCancellationRequested(); int count = stream.Read(output, copied, output.Length - copied); if (count == 0) throw new InvalidDataException("Native Access long value is truncated."); copied += count; }
        return output;
    }
}
