namespace OfficeIMO.Publisher.Internal;

/// <summary>Replaces disjoint ranges addressed in the immutable source stream, retaining every intervening byte.</summary>
internal sealed record PublisherNativeStreamEdit(int Offset, int Length, byte[] Bytes) {
    internal static byte[] Apply(byte[] source, IReadOnlyList<PublisherNativeStreamEdit> edits, long maximumBytes, CancellationToken token) {
        var ordered = edits.OrderBy(edit => edit.Offset).ToArray();
        long size = source.Length; int end = 0;
        foreach (var edit in ordered) {
            token.ThrowIfCancellationRequested();
            if (edit.Offset < end || edit.Offset < 0 || edit.Length < 0 || edit.Offset > source.Length - edit.Length)
                throw new InvalidDataException("Overlapping or invalid native source edits.");
            end = edit.Offset + edit.Length; size = checked(size + edit.Bytes.Length - edit.Length);
        }
        if (size < 0 || size > maximumBytes || size > int.MaxValue) throw new InvalidDataException("Publisher native output byte limit exceeded.");
        var result = new byte[(int)size]; int input = 0, output = 0;
        foreach (var edit in ordered) {
            token.ThrowIfCancellationRequested();
            int retained = edit.Offset - input;
            Buffer.BlockCopy(source, input, result, output, retained); output += retained;
            Buffer.BlockCopy(edit.Bytes, 0, result, output, edit.Bytes.Length); output += edit.Bytes.Length;
            input = edit.Offset + edit.Length;
        }
        Buffer.BlockCopy(source, input, result, output, source.Length - input);
        return result;
    }
}
