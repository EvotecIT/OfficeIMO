using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordSettings {
    /// <summary>Gets or sets the document's decimal separator, including decimal tab alignment.</summary>
    /// <remarks>A non-null value must be one nonsurrogate, noncontrol character. Null removes the explicit setting.</remarks>
    public string? DecimalSymbol {
        get => _document._wordprocessingDocument.MainDocumentPart?.DocumentSettingsPart?.Settings?
            .GetFirstChild<DecimalSymbol>()?.Val?.Value;
        set {
            if (value != null && (value.Length != 1 || char.IsControl(value[0]) || char.IsSurrogate(value[0])))
                throw new ArgumentException("A decimal symbol must be one noncontrol Unicode character.", nameof(value));
            var settings = _document._wordprocessingDocument.MainDocumentPart?.DocumentSettingsPart?.Settings;
            if (settings == null) return;
            var symbol = settings.GetFirstChild<DecimalSymbol>();
            if (value == null) { symbol?.Remove(); return; }
            if (symbol == null) { symbol = new DecimalSymbol(); settings.AddChild(symbol, true); }
            symbol.Val = value;
        }
    }
}
