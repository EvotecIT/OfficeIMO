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
}
