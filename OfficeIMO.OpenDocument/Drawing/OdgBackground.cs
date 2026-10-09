using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

/// <summary>XML-backed background paint in a page or master's drawing-page style chain.</summary>
/// <remarks>Page declarations override master paint except for native no-fill, which retains the master.
/// Master edits affect all referencing pages. Referenced gradient definitions and image bytes remain shared.</remarks>
public sealed class OdgBackground {
    private static readonly XName PropertiesName = OdfNamespaces.Style + "drawing-page-properties";
    private readonly OdgDocument _document;
    private readonly XElement _owner;
    private readonly string _partPath;
    internal OdgBackground(OdgDocument document, XElement owner, string partPath) {
        _document = document; _owner = owner; _partPath = partPath;
    }

    /// <summary>Resolved native fill token in this owner's style chain, or null when absent. It does not include the other owner's paint.</summary>
    public string? FillMode => ReadProperty(OdfNamespaces.Draw + "fill");

    /// <summary>Resolved solid fill color, or null for another fill mode. Setting null selects native no-fill.</summary>
    /// <remarks>Native page no-fill retains master paint; master no-fill removes master paint.</remarks>
    public OdfColor? FillColor {
        get => FillMode == "solid" ? OdfColor.Parse(ReadProperty(OdfNamespaces.Draw + "fill-color")
            ?? throw new InvalidDataException("A solid background requires a resolved fill color.")) : null;
        set {
            OdfStyle style = EnsureStyle();
            style.SetProperty(PropertiesName, OdfNamespaces.Draw + "fill", value.HasValue ? "solid" : "none");
            style.SetProperty(PropertiesName, OdfNamespaces.Draw + "fill-color", value?.ToString());
        }
    }

    /// <summary>Resolved gradient binding, or null for another fill mode. Setting null selects native no-fill.</summary>
    /// <remarks>Uses the common native two-color model. Projection has a narrower, reported profile; definitions remain shared.</remarks>
    public string? FillGradientName {
        get => FillMode == "gradient" ? ReadProperty(OdfNamespaces.Draw + "fill-gradient-name")
            ?? throw new InvalidDataException("A gradient background requires a named definition.") : null;
        set {
            if (value != null) {
                OdfStyleRepository.ValidateStyleName(value);
                OdfGradient gradient = _document.Styles.FindGradient(value) ?? throw new ArgumentException("Unknown gradient '" + value + "'.", nameof(value));
                _ = gradient.Pattern;
            }
            OdfStyle style = EnsureStyle();
            style.SetProperty(PropertiesName, OdfNamespaces.Draw + "fill", value == null ? "none" : "gradient");
            style.SetProperty(PropertiesName, OdfNamespaces.Draw + "fill-gradient-name", value);
        }
    }

    /// <summary>Uniform opacity from zero to one in this owner's style chain. Null removes the local override.</summary>
    /// <remarks>Opacity gradients remain preserved and cannot be represented or replaced by this property.</remarks>
    public double? FillOpacity {
        get { EnsureUniformOpacity(); string? value = ReadProperty(OdfNamespaces.Draw + "opacity"); return value == null ? null : OdfOpacity.Parse(value); }
        set { OdfOpacity.Validate(value); EnsureUniformOpacity(); WriteProperty(OdfNamespaces.Draw + "opacity", OdfOpacity.Format(value)); }
    }

    /// <summary>Resolved native gradient band count. Zero means automatic; three or more means fixed bands. Null removes the local override.</summary>
    /// <remarks>Fixed bands are preserved but are outside the current projection profile.</remarks>
    public int? GradientStepCount {
        get {
            string? value = ReadProperty(OdfNamespaces.Draw + "gradient-step-count");
            if (value == null) return null;
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int count) || count < 0 || count is 1 or 2)
                throw new InvalidDataException("Invalid background gradient step count.");
            return count;
        }
        set {
            if (value.HasValue && (value.Value < 0 || value.Value is 1 or 2)) throw new ArgumentOutOfRangeException(nameof(value));
            WriteProperty(OdfNamespaces.Draw + "gradient-step-count", value?.ToString(CultureInfo.InvariantCulture));
        }
    }

    /// <summary>Resolved bitmap definition name, or null for another fill mode.</summary>
    public string? FillImageName => FillMode == "bitmap" ? ReadProperty(OdfNamespaces.Draw + "fill-image-name")
        ?? throw new InvalidDataException("A bitmap background requires a named definition.") : null;

    /// <summary>Returns embedded bitmap source bytes without fetching linked content.</summary>
    public byte[] GetBitmapBytes() => _document.Styles.ReadFillImage(FillImageName
        ?? throw new InvalidOperationException("Background paint is not a bitmap."));

    /// <summary>Embeds a supported raster and selects stretched bitmap fill without changing opacity or paint area.</summary>
    /// <remarks>Uses the shared PNG/JPEG/GIF/BMP/TIFF/WebP drawing/PDF resource profile. Image entries are deduplicated;
    /// a new common definition prevents changes to other bindings. Inactive paint properties and unrelated XML are retained.</remarks>
    public void SetBitmap(byte[] data, string fileName = "background.png") {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (!OfficeImageReader.TryValidateContent(data, fileName, out OfficeImageInfo info) || !OdgPage.IsWithinRasterProjectionProfile(info))
            throw new ArgumentException("Background data must be a complete raster inside the shared drawing/PDF profile.", nameof(data));
        _ = GetStyle(); // Reject an unresolved binding before embedding bytes.
        string path = OdfImageStore.Add(_document, data, fileName);
        string name = _document.Styles.CreateFillImage(path);
        OdfStyle style = EnsureStyle();
        style.SetProperty(PropertiesName, OdfNamespaces.Draw + "fill", "bitmap");
        style.SetProperty(PropertiesName, OdfNamespaces.Draw + "fill-image-name", name);
        style.SetProperty(PropertiesName, OdfNamespaces.Style + "repeat", "stretch");
    }

    /// <summary>Selects native no-fill, retaining inactive declarations and resource definitions.</summary>
    /// <remarks>On a page this retains master paint. On a master this removes its background paint.</remarks>
    public void UseNoFill() => WriteProperty(OdfNamespaces.Draw + "fill", "none");

    internal string? ReadProperty(XName name) => _document.Styles.ResolveWithDefault(GetStyle(), OdfStyleFamily.DrawingPage)
        .Select(style => (string?)style.Element.Element(PropertiesName)?.Attribute(name)).FirstOrDefault(value => value != null);
    internal void WriteProperty(XName name, string? value) => EnsureStyle().SetProperty(PropertiesName, name, value);
    private OdfStyle EnsureStyle() {
        _ = GetStyle();
        return _document.Styles.EnsureAutomaticStyle(_owner, OdfNamespaces.Draw + "style-name", OdfStyleFamily.DrawingPage, "ofBg", _partPath);
    }
    private OdfStyle? GetStyle() {
        string? name = (string?)_owner.Attribute(OdfNamespaces.Draw + "style-name");
        if (name == null) return null;
        OdfStyleRepository.ValidateStyleName(name);
        return _document.Styles.FindInPart(OdfStyleFamily.DrawingPage, name, _partPath)
            ?? throw new InvalidDataException("Unresolved drawing-page background style '" + name + "'.");
    }
    private void EnsureUniformOpacity() {
        if (!string.IsNullOrEmpty(ReadProperty(OdfNamespaces.Draw + "opacity-name")))
            throw new NotSupportedException("Uniform background opacity cannot represent or replace an opacity gradient.");
    }
}
