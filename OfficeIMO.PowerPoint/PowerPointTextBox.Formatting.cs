namespace OfficeIMO.PowerPoint;

public partial class PowerPointTextBox {
    /// <summary>Sets <see cref="Bold"/> for the entire text box and returns this instance.</summary>
    public PowerPointTextBox SetBold(bool bold = true) {
        Bold = bold;
        return this;
    }

    /// <summary>Sets <see cref="Italic"/> for the entire text box and returns this instance.</summary>
    public PowerPointTextBox SetItalic(bool italic = true) {
        Italic = italic;
        return this;
    }

    /// <summary>Sets <see cref="FontSize"/> for the entire text box and returns this instance.</summary>
    public PowerPointTextBox SetFontSize(OfficeLength? fontSize) {
        FontSize = fontSize;
        return this;
    }

    /// <summary>Sets <see cref="FontFamily"/> for the entire text box and returns this instance.</summary>
    public PowerPointTextBox SetFontFamily(string? fontFamily) {
        FontFamily = fontFamily;
        return this;
    }

    /// <summary>Sets <see cref="Color"/> for the entire text box and returns this instance.</summary>
    public PowerPointTextBox SetColor(OfficeColor? color) {
        Color = color;
        return this;
    }

}
