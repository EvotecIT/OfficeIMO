namespace OfficeIMO.IWork.Internal;

internal sealed partial class IWorkTableNumberFormatCatalog {
    // Pinned independent FormatStructArchive qualifies day-only and hour/minute
    // ranges; the latter also has paired Numbers 14.5 export evidence. Do not infer
    // defaults or discard custom, automatic, scaling or control metadata.
    private static IWorkNumberFormat? ReadDurationFormat(IWorkWireMessage message) {
        int[] fields = { 1, 7, 15, 16, 40 };
        if (message.TotalFieldCount != fields.Length || fields.Any(field =>
                message.FieldCount(field) != 1 || message.HasUnexpectedWireKind(field, IWorkWireKind.Varint))
            || message.GetUnsigned(1) != 268 || message.GetUnsigned(7) != 1
            || message.GetUnsigned(40) != 0) return null;
        ulong? largest = message.GetUnsigned(15), smallest = message.GetUnsigned(16);
        if (!(largest == 4 && smallest == 8 || largest == 2 && smallest == 2)) return null;
        return new IWorkNumberFormat(IWorkNumberFormatKind.Duration, null, false,
            IWorkNegativeNumberStyle.Minus, durationFormat: new IWorkDurationFormat(
                (IWorkDurationUnit)largest.Value, (IWorkDurationUnit)smallest.Value));
    }
}
