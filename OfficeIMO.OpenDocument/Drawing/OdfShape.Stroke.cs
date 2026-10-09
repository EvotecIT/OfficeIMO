using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public abstract partial class OdfShape {
    /// <summary>Inherited stroke cap; null removes the local override and may restore an inherited value.</summary>
    public OfficeStrokeLineCap? StrokeLineCap {
        get => ReadGraphicProperty(OdfNamespaces.Svg + "stroke-linecap") switch {
            null => null, "butt" => OfficeStrokeLineCap.Butt, "round" => OfficeStrokeLineCap.Round, "square" => OfficeStrokeLineCap.Square,
            _ => throw new NotSupportedException("Unsupported ODF stroke cap.")
        };
        set {
            if (value.HasValue && !Enum.IsDefined(typeof(OfficeStrokeLineCap), value.Value)) throw new ArgumentOutOfRangeException(nameof(value));
            EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Svg + "stroke-linecap", value?.ToString().ToLowerInvariant());
        }
    }
    /// <summary>Inherited miter, round, or bevel join; null removes the local override.</summary>
    /// <remarks>Native middle and none values remain preserved but are outside the shared rendering profile.</remarks>
    public OfficeStrokeLineJoin? StrokeLineJoin {
        get => ReadGraphicProperty(OdfNamespaces.Draw + "stroke-linejoin") switch {
            null => null, "miter" => OfficeStrokeLineJoin.Miter, "round" => OfficeStrokeLineJoin.Round, "bevel" => OfficeStrokeLineJoin.Bevel,
            _ => throw new NotSupportedException("Unsupported ODF stroke join.")
        };
        set {
            if (value.HasValue && !Enum.IsDefined(typeof(OfficeStrokeLineJoin), value.Value)) throw new ArgumentOutOfRangeException(nameof(value));
            EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "stroke-linejoin", value?.ToString().ToLowerInvariant());
        }
    }
    /// <summary>Effective named dash pattern when the stroke is dashed. Setting null explicitly selects a solid stroke.</summary>
    /// <remarks>Assigning a name enables a dashed stroke without changing its color or width. Definitions are shared; shape styles use copy-on-write.</remarks>
    public string? StrokeDashName {
        get {
            if (ReadGraphicProperty(OdfNamespaces.Draw + "stroke") != "dash") return null;
            string? name = ReadGraphicProperty(OdfNamespaces.Draw + "stroke-dash");
            if (string.IsNullOrWhiteSpace(name)) throw new InvalidDataException("Dashed stroke has no dash definition reference.");
            return name;
        }
        set {
            if (value != null) {
                OdfStyleRepository.ValidateStyleName(value);
                OdfStrokeDash dash = Document.Styles.FindStrokeDash(value) ?? throw new ArgumentException("Unknown stroke-dash '" + value + "'.", nameof(value));
                _ = dash.Pattern;
            }
            OdfStyle style = EnsureGraphicStyle();
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "stroke", value == null ? "solid" : "dash");
            style.SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "stroke-dash", value);
        }
    }
    internal bool HasStackedStrokeDashes => !string.IsNullOrWhiteSpace(ReadGraphicProperty(OdfNamespaces.Draw + "stroke-dash-names"));
    internal OdfStrokeDashPattern? ResolveStrokeDash() {
        string? name = StrokeDashName;
        return name == null ? null : (Document.Styles.FindStrokeDash(name) ?? throw new InvalidDataException("Missing stroke-dash '" + name + "'.")).Pattern;
    }
}
