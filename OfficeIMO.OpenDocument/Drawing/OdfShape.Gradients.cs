namespace OfficeIMO.OpenDocument;

public abstract partial class OdfShape {
    /// <summary>Effective named gradient when the fill mode is gradient. Setting null explicitly disables the fill.</summary>
    /// <remarks>Definitions are shared. Shape edits use copy-on-write and retain inactive solid colors and unrelated style properties.</remarks>
    public string? FillGradientName {
        get {
            if (!HasGradientFill) return null;
            string? name = ReadGraphicProperty(OdfNamespaces.Draw + "fill-gradient-name");
            if (string.IsNullOrWhiteSpace(name)) throw new InvalidDataException("Gradient fill has no definition reference.");
            return name;
        }
        set {
            if (value != null) {
                OdfStyleRepository.ValidateStyleName(value);
                OdfGradient gradient = Document.Styles.FindGradient(value) ?? throw new ArgumentException("Unknown gradient '" + value + "'.", nameof(value));
                _ = gradient.Pattern;
            }
            OdfStyle style = EnsureGraphicStyle();
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill", value == null ? "none" : "gradient");
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "fill-gradient-name", value);
        }
    }
    /// <summary>Inherited gradient step count: zero selects automatic interpolation, three or more select fixed bands. Null removes the local override.</summary>
    /// <remarks>Fixed bands remain preserved but are outside the current drawing projection profile.</remarks>
    public int? GradientStepCount {
        get {
            string? value = ReadGraphicProperty(OdfNamespaces.Draw + "gradient-step-count");
            if (value == null) return null;
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count) || count < 0 || count is 1 or 2)
                throw new InvalidDataException("Invalid gradient step count.");
            return count;
        }
        set {
            if (value.HasValue && (value.Value < 0 || value.Value is 1 or 2)) throw new ArgumentOutOfRangeException(nameof(value));
            EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "gradient-step-count", value?.ToString(CultureInfo.InvariantCulture));
        }
    }
    internal bool HasGradientFill => ReadGraphicProperty(OdfNamespaces.Draw + "fill") == "gradient";
    internal OdfGradient ResolveFillGradient() => Document.Styles.FindGradient(FillGradientName!) ?? throw new InvalidDataException("Missing gradient '" + FillGradientName + "'.");
}
