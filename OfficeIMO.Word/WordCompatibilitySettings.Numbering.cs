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
        get => ReadOption("doNotUseIndentAsNumberingTabStop");
        set => SetOption("doNotUseIndentAsNumberingTabStop", value);
    }
}
