namespace OfficeIMO.PowerPoint;

public partial class PowerPointTableCell {
    /// <summary>Sets <see cref="Bold"/> for the entire cell and returns this instance.</summary>
    public PowerPointTableCell SetBold(bool bold = true) {
        Bold = bold;
        return this;
    }

    /// <summary>Sets <see cref="Italic"/> for the entire cell and returns this instance.</summary>
    public PowerPointTableCell SetItalic(bool italic = true) {
        Italic = italic;
        return this;
    }

    /// <summary>Sets <see cref="FontSize"/> for the entire cell and returns this instance.</summary>
    public PowerPointTableCell SetFontSize(OfficeLength? fontSize) {
        FontSize = fontSize;
        return this;
    }

    /// <summary>Sets <see cref="FontFamily"/> for the entire cell and returns this instance.</summary>
    public PowerPointTableCell SetFontFamily(string? fontFamily) {
        FontFamily = fontFamily;
        return this;
    }

    /// <summary>Sets <see cref="Color"/> for the entire cell and returns this instance.</summary>
    public PowerPointTableCell SetColor(OfficeColor? color) {
        Color = color;
        return this;
    }

}
