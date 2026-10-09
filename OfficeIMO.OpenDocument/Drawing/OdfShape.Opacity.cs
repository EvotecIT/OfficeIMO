namespace OfficeIMO.OpenDocument;

public abstract partial class OdfShape {
    /// <summary>Uniform fill opacity from zero to one, including inherited values. Null means no declared value.</summary>
    /// <remarks>Assigning null removes the local override. Opacity gradients require a separate rendering profile and are rejected by this property.</remarks>
    public double? FillOpacity {
        get {
            EnsureUniformFillOpacity();
            return ReadOpacity(OdfNamespaces.Draw + "opacity", allowFraction: false);
        }
        set {
            OdfOpacity.Validate(value);
            EnsureUniformFillOpacity();
            WriteOpacity(OdfNamespaces.Draw + "opacity", value);
        }
    }

    /// <summary>Stroke opacity from zero to one, including inherited values. Null means no declared value; assignment removes the local override.</summary>
    public double? StrokeOpacity {
        get => ReadOpacity(OdfNamespaces.Svg + "stroke-opacity", allowFraction: true);
        set { OdfOpacity.Validate(value); WriteOpacity(OdfNamespaces.Svg + "stroke-opacity", value); }
    }

    /// <summary>Image-pixel opacity from zero to one, including inherited values. This is separate from the frame's fill opacity.</summary>
    /// <remarks>Null means no declared value; assignment removes the local override.</remarks>
    public double? ImageOpacity {
        get => ReadOpacity(OdfNamespaces.Draw + "image-opacity", allowFraction: false);
        set { OdfOpacity.Validate(value); WriteOpacity(OdfNamespaces.Draw + "image-opacity", value); }
    }

    internal bool HasOpacityGradient => !string.IsNullOrEmpty(ReadGraphicProperty(OdfNamespaces.Draw + "opacity-name"));
    private void EnsureUniformFillOpacity() {
        if (HasOpacityGradient) throw new NotSupportedException("Uniform fill opacity cannot represent or replace an opacity gradient.");
    }
    private double? ReadOpacity(XName name, bool allowFraction) {
        string? value = ReadGraphicProperty(name);
        if (value == null) return null;
        return OdfOpacity.Parse(value, allowFraction);
    }
    private void WriteOpacity(XName name, double? value) => EnsureGraphicStyle().SetProperty(
        OdfNamespaces.Style + "graphic-properties", name, OdfOpacity.Format(value));
}
