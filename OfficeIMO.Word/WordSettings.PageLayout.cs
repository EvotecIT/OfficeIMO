using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordSettings {
    /// <summary>Gets or sets whether the document gutter is placed at the top of each page.</summary>
    public bool GutterAtTop {
        get => ReadSetting("gutterAtTop").IsOn();
        set => SetSettingAttribute("gutterAtTop", "val", value ? "true" : "false");
    }

    /// <summary>Gets or sets whether odd and even pages use mirrored inside/outside margins.</summary>
    public bool MirrorMargins {
        get => ReadSetting("mirrorMargins").IsOn();
        set {
            if (value) SetSettingAttribute("mirrorMargins", "val", "true");
            else RemoveSetting("mirrorMargins");
        }
    }
}
