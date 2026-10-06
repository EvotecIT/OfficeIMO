namespace OfficeIMO.Pdf;

/// <summary>Reuses a horizontal embedded CID font's existing Unicode mapping and PDF widths.</summary>
internal sealed class PdfSourceTextFont {
    private readonly ToUnicodeCMap _cmap;
    private readonly Func<byte[], double> _width;
    private readonly int _codeHexLength;

    internal PdfSourceTextFont(PdfFontResource font, Func<byte[], double> width) {
        ResourceName = font.ResourceName;
        BaseFont = font.BaseFont;
        _cmap = font.CMap!;
        _width = width;
        _codeHexLength = font.FontSubtype == "Type0" ? 4 : 2;
    }

    internal string ResourceName { get; }
    internal string BaseFont { get; }

    internal string Encode(string text) {
        if (!_cmap.TryEncodeTextCodes(text, out IReadOnlyList<string> codes) ||
            codes.Any(code => code.Length != _codeHexLength)) {
            throw new NotSupportedException("The source font '" + BaseFont +
                "' cannot encode this text with its existing subset. Choose an explicit Font to permit substitution.");
        }
        return string.Concat(codes);
    }

    internal double Measure(string text, double fontSize) {
        string hex = Encode(text);
        var bytes = new byte[hex.Length / 2];
        for (int index = 0; index < bytes.Length; index++) {
            bytes[index] = Convert.ToByte(hex.Substring(index * 2, 2), 16);
        }
        return _width(bytes) * fontSize / 1000D;
    }
}
