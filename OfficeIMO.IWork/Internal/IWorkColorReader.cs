namespace OfficeIMO.IWork.Internal;

/// <summary>Decodes the common native normalized RGB/gray color representation.</summary>
internal static class IWorkColorReader {
    internal static bool TryRead(IWorkWireMessage owner, int field, out IWorkColor? color,
        ref bool complete) {
        color = null;
        bool hasColor = owner.HasField(field);
        IWorkWireMessage? message = IWorkObjectIndex.TryGetMessage(owner, field, out bool malformedColor);
        if (owner.FieldCount(field) > 1 || owner.HasUnexpectedWireKind(field, IWorkWireKind.Bytes)
            || malformedColor || hasColor && message == null) {
            complete = false;
            return false;
        }
        if (message == null) return false;
        bool hasWhite = message.HasField(11);
        bool hasAnyRgb = message.HasField(3) || message.HasField(4) || message.HasField(5);
        bool hasCompleteRgb = message.HasField(3) && message.HasField(4) && message.HasField(5);
        float? white = message.GetFloat(11);
        float red = white ?? message.GetFloat(3) ?? 0;
        float green = white ?? message.GetFloat(4) ?? 0;
        float blue = white ?? message.GetFloat(5) ?? 0;
        float alpha = message.GetFloat(6) ?? 1;
        if (new[] { 3, 4, 5, 6, 11 }.Any(component =>
                message.FieldCount(component) > 1
                || message.HasUnexpectedWireKind(component, IWorkWireKind.Fixed32)
                || message.HasField(component) && !message.GetFloat(component).HasValue)
            || hasWhite == hasAnyRgb || hasAnyRgb && !hasCompleteRgb
            || !new[] { red, green, blue, alpha }.All(value =>
                !float.IsNaN(value) && !float.IsInfinity(value) && value >= 0f && value <= 1f)) {
            complete = false;
            return false;
        }
        color = new IWorkColor(Component(red), Component(green), Component(blue),
            alpha >= 1f ? byte.MaxValue : (byte)Math.Floor(alpha * byte.MaxValue));
        return true;
    }

    private static byte Component(float value) =>
        (byte)Math.Round(value * 255, MidpointRounding.AwayFromZero);

    /// <summary>Qualifies the opaque RGB/gray archive shape shared by fills and table strokes.</summary>
    internal static bool IsSupportedColor(IWorkWireMessage color) {
        int[] fields = { 1, 3, 4, 5, 6, 11, 12, 13 };
        if (color.TotalFieldCount != fields.Sum(color.FieldCount)) return false;
        foreach (int field in new[] { 1, 12 })
            if (color.FieldCount(field) > 1 || color.HasUnexpectedWireKind(field, IWorkWireKind.Varint)) return false;
        if (color.HasField(12) && color.GetUnsigned(12) != 1) return false;
        // Keynote 15.4's additional opaque-sRGB scalar is qualified only at 1.
        if (color.HasField(13) && (color.FieldCount(13) != 1
            || color.HasUnexpectedWireKind(13, IWorkWireKind.Fixed32)
            || color.GetFloat(13) != 1f || color.GetUnsigned(1) != 1
            || color.GetUnsigned(12) != 1 || color.HasField(11))) return false;
        return !color.HasField(1) || color.GetUnsigned(1) == (color.HasField(11) ? 3UL : 1UL);
    }
}
