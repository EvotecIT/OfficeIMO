using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordCompatibilitySettings {
    /// <summary>
    /// Gets or sets whether tab-suffixed numbering uses a custom or default tab
    /// instead of an implicit stop at the paragraph's hanging indent.
    /// </summary>
    /// <remarks>
    /// The DOCX default is false. Word 2013 layout ignores the stored option.
    /// Explicit enabled and disabled values survive DOCX editing and saving.
    /// Native DOC import projects the effective enabled layout. Native DOC saving
    /// cannot retain an explicitly disabled value; save as DOCX to retain it.
    /// </remarks>
    public bool DoNotUseIndentAsNumberingTabStop {
        get {
            DoNotUseIndentAsNumberingTabStop? setting = _wordprocessingDocument.MainDocumentPart?
                .DocumentSettingsPart?.Settings?.GetFirstChild<Compatibility>()?
                .GetFirstChild<DoNotUseIndentAsNumberingTabStop>();
            return setting != null && (setting.Val?.Value ?? true);
        }
        set {
            Compatibility? compatibility = _wordprocessingDocument.MainDocumentPart?
                .DocumentSettingsPart?.Settings?.GetFirstChild<Compatibility>();
            if (compatibility == null) {
                compatibility = new Compatibility();
                GetSettings().AddChild(compatibility, true);
            }
            DoNotUseIndentAsNumberingTabStop? setting = compatibility.GetFirstChild<DoNotUseIndentAsNumberingTabStop>();
            if (setting == null) {
                setting = new DoNotUseIndentAsNumberingTabStop();
                compatibility.AddChild(setting, true);
            }
            setting.Val = value;
        }
    }
}
