using OfficeIMO.Drawing;

namespace OfficeIMO.IWork.Internal;

/// <summary>One declared stroke, including a rejected declaration or an explicit invisible edge.</summary>
internal sealed class IWorkBorderDeclaration(OfficeBorderSide? side, bool isSupported) {
    internal OfficeBorderSide? Side { get; } = side;
    internal bool IsSupported { get; } = isSupported;
    internal static IWorkBorderDeclaration Invalid { get; } = new(null, false);
    internal static IWorkBorderDeclaration None { get; } = new(new OfficeBorderSide(OfficeColor.Black, 0), true);
}

/// <summary>Decodes solid and empty native table strokes without reinterpreting patterned or transparent strokes.</summary>
internal static class IWorkStrokeReader {
    internal static IWorkBorderDeclaration Read(IWorkArchiveRecord owner, IWorkWireMessage message,
        int field, string path, IWorkSourceReferenceIssueCollector references, bool deprecated = false) {
        IWorkWireMessage? stroke = IWorkObjectIndex.TryGetMessage(message, field, out bool malformed);
        bool valid = !malformed && message.FieldCount(field) == 1
            && !message.HasUnexpectedWireKind(field, IWorkWireKind.Bytes) && stroke != null;
        if (valid && deprecated) {
            valid = stroke!.TotalFieldCount == Enumerable.Range(1, 4).Sum(stroke.FieldCount)
                && ValidScalar(stroke, 2, IWorkWireKind.Varint) && stroke.GetUnsigned(2).GetValueOrDefault() == 0
                && ValidScalar(stroke, 3, IWorkWireKind.Fixed32) && (!stroke.HasField(3) || stroke.GetFloat(3) == 1)
                && ValidScalar(stroke, 4, IWorkWireKind.Varint) && stroke.GetUnsigned(4).GetValueOrDefault() <= 1;
            if (valid && stroke.GetUnsigned(4) == 1) return IWorkBorderDeclaration.None;
            IWorkWireMessage? nested = IWorkObjectIndex.TryGetMessage(stroke, 1, out bool badNested);
            valid &= !badNested && stroke.FieldCount(1) == 1
                && !stroke.HasUnexpectedWireKind(1, IWorkWireKind.Bytes) && nested != null;
            stroke = nested;
        }
        OfficeBorderSide? side = valid ? Decode(stroke!, ref valid) : null;
        if (!valid) {
            references.Declarations.Record(owner, path, message.FieldCount(field), IWorkSourceDeclarationIssueKind.InvalidValue);
            return IWorkBorderDeclaration.Invalid;
        }
        return new IWorkBorderDeclaration(side, true);
    }

    private static OfficeBorderSide? Decode(IWorkWireMessage stroke, ref bool valid) {
        valid = stroke.TotalFieldCount == Enumerable.Range(1, 6).Sum(stroke.FieldCount)
            && ValidScalar(stroke, 2, IWorkWireKind.Fixed32)
            && ValidScalar(stroke, 3, IWorkWireKind.Varint) && stroke.GetUnsigned(3).GetValueOrDefault() == 0
            && ValidScalar(stroke, 4, IWorkWireKind.Varint) && stroke.GetUnsigned(4).GetValueOrDefault() == 0
            && ValidScalar(stroke, 5, IWorkWireKind.Fixed32)
            && (!stroke.HasField(5) || stroke.GetFloat(5) is float miter && Finite(miter) && miter > 0);
        IWorkWireMessage? pattern = IWorkObjectIndex.TryGetMessage(stroke, 6, out bool malformed);
        valid &= !malformed && stroke.FieldCount(6) <= 1 && !stroke.HasUnexpectedWireKind(6, IWorkWireKind.Bytes)
            && (!stroke.HasField(6) || pattern != null);
        ulong kind = 1;
        if (pattern != null) {
            kind = pattern.GetUnsigned(1).GetValueOrDefault();
            valid &= pattern.TotalFieldCount == Enumerable.Range(1, 4).Sum(pattern.FieldCount)
                && ValidScalar(pattern, 1, IWorkWireKind.Varint) && kind is 1 or 2
                && ValidScalar(pattern, 2, IWorkWireKind.Fixed32) && pattern.GetFloat(2).GetValueOrDefault() == 0
                && ValidScalar(pattern, 3, IWorkWireKind.Varint) && pattern.GetUnsigned(3).GetValueOrDefault() == 0
                && pattern.FieldCount(4) <= 6 // Native solid/empty padding has six zero entries; do not scan an unbounded unused pattern.
                && !pattern.HasUnexpectedWireKind(4, IWorkWireKind.Fixed32)
                && pattern.GetRepeatedFloat(4).All(value => value == 0);
        }
        float width = stroke.GetFloat(2).GetValueOrDefault();
        valid &= Finite(width) && width >= 0;
        if (!valid) return null;
        if (kind == 2 || width == 0) return IWorkBorderDeclaration.None.Side;
        IWorkColor? color;
        valid = IWorkColorReader.TryRead(stroke, 1, out color, ref valid)
            && color is { Alpha: byte.MaxValue }
            && IWorkColorReader.IsSupportedColor(IWorkObjectIndex.TryGetMessage(stroke, 1)!);
        return valid ? new OfficeBorderSide(OfficeColor.FromRgb(color!.Red, color.Green, color.Blue), width) : null;
    }

    internal static bool ValidScalar(IWorkWireMessage message, int field, IWorkWireKind wire) =>
        message.FieldCount(field) <= 1 && !message.HasUnexpectedWireKind(field, wire);

    private static bool Finite(float value) => !float.IsNaN(value) && !float.IsInfinity(value);
}
