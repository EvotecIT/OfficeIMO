using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordSettings {
    /// <summary>Gets or sets the document's decimal separator, including decimal tab alignment.</summary>
    /// <remarks>A non-null value must be one nonsurrogate, noncontrol character. Null removes the explicit setting.</remarks>
    public string? DecimalSymbol {
        get => ReadSetting("decimalSymbol").WVal();
        set {
            if (value == null) {
                RemoveSetting("decimalSymbol");
                return;
            }
            if (value.Length != 1 || char.IsControl(value[0]) || char.IsSurrogate(value[0]))
                throw new ArgumentException("A decimal symbol must be one noncontrol Unicode character.", nameof(value));
            SetSettingAttribute("decimalSymbol", "val", value);
        }
    }
}
