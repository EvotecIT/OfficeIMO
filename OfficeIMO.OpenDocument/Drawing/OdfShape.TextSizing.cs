using System.Text.RegularExpressions;

namespace OfficeIMO.OpenDocument;

public abstract partial class OdfShape {
    /// <summary>Inherited native text-box height growth; null removes the local declaration.</summary>
    /// <remarks>An omitted value remains unspecified; producer defaults can differ. This flag is independent of width growth and text fitting.</remarks>
    /// <exception cref="NotSupportedException">The effective declaration is not an ODF boolean.</exception>
    public bool? AutoGrowHeight {
        get => ReadTextGrowth("auto-grow-height");
        set => SetTextGrowth("auto-grow-height", value);
    }

    /// <summary>Inherited native text-box width growth; null removes the local declaration.</summary>
    /// <remarks>Draw projection approximates explicitly declared growth for horizontal top-aligned ordinary text boxes. With both axes enabled, width grows before height; saved ODF geometry remains unchanged.</remarks>
    /// <exception cref="NotSupportedException">The effective declaration is not an ODF boolean.</exception>
    public bool? AutoGrowWidth {
        get => ReadTextGrowth("auto-grow-width");
        set => SetTextGrowth("auto-grow-width", value);
    }

    /// <summary>Instance minimum height on the direct text-box container; null removes the attribute.</summary>
    /// <remarks>This is distinct from the graphic-style creation minimum. Editing preserves native lengths and percentages without evaluating pairs. Draw projection evaluates absolute instance minima and matching-unit maxima in its ordinary text-box profile.</remarks>
    public OdfLength? TextBoxMinimumHeight {
        get => ReadTextBoxSize("min-height");
        set => SetTextBoxSize("min-height", value);
    }

    /// <summary>Instance minimum width on the direct text-box container; null removes the attribute.</summary>
    public OdfLength? TextBoxMinimumWidth {
        get => ReadTextBoxSize("min-width");
        set => SetTextBoxSize("min-width", value);
    }

    /// <summary>Instance maximum height on the direct text-box container; null removes the attribute.</summary>
    public OdfLength? TextBoxMaximumHeight {
        get => ReadTextBoxSize("max-height");
        set => SetTextBoxSize("max-height", value);
    }

    /// <summary>Instance maximum width on the direct text-box container; null removes the attribute.</summary>
    public OdfLength? TextBoxMaximumWidth {
        get => ReadTextBoxSize("max-width");
        set => SetTextBoxSize("max-width", value);
    }

    private bool? ReadTextGrowth(string name) => ReadGraphicProperty(OdfNamespaces.Draw + name) switch {
        null => null, "true" => true, "false" => false,
        _ => throw new NotSupportedException("Unsupported ODF text-box growth declaration.")
    };

    private void SetTextGrowth(string name, bool? value) => EnsureGraphicStyle().SetProperty(
        OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + name,
        value.HasValue ? value.Value ? "true" : "false" : null);

    private OdfLength? ReadTextBoxSize(string name) {
        string? value = (string?)RequireTextBoxSizeContainer().Attribute(OdfNamespaces.Fo + name);
        return value == null ? null : OdfLength.Parse(value);
    }

    private void SetTextBoxSize(string name, OdfLength? value) {
        XElement container = RequireTextBoxSizeContainer();
        if (value.HasValue) {
            string lexical = value.Value.ToString();
            if (!Regex.IsMatch(lexical, @"\A(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:cm|mm|in|pt|pc|px|%)\z", RegexOptions.CultureInvariant) ||
                !double.TryParse(lexical.Substring(0, lexical.Length - (lexical.EndsWith("%", StringComparison.Ordinal) ? 1 : 2)),
                    NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out double magnitude) ||
                double.IsNaN(magnitude) || double.IsInfinity(magnitude))
                throw new ArgumentException("Text-box sizes must use finite nonnegative ODF lengths or percentages.", nameof(value));
        }
        container.SetAttributeValue(OdfNamespaces.Fo + name, value?.ToString());
        Dirty();
    }

    private XElement RequireTextBoxSizeContainer() => Element.Name == OdfNamespaces.Draw + "frame" &&
        Element.Element(OdfNamespaces.Draw + "text-box") is { } container ? container :
        throw new NotSupportedException("Instance text-box sizes require a direct draw:text-box inside a frame.");
}
