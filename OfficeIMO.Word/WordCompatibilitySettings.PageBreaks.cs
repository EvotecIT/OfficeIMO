using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordCompatibilitySettings {
    /// <summary>
    /// Gets or sets whether a manual page break at the end of a paragraph moves
    /// its paragraph mark onto the first line of the following page.
    /// </summary>
    /// <remarks>
    /// The DOCX default is false. Word 2013 layout ignores the stored option.
    /// Both enabled and disabled values are retained when explicitly assigned.
    /// Native DOC import uses the legacy enabled layout. Saving an explicitly
    /// disabled value to native DOC is unsupported; save as DOCX to retain it.
    /// </remarks>
    public bool SplitPageBreakAndParagraphMark {
        get {
            SplitPageBreakAndParagraphMark? setting = _wordprocessingDocument.MainDocumentPart?
                .DocumentSettingsPart?.Settings?.GetFirstChild<Compatibility>()?
                .GetFirstChild<SplitPageBreakAndParagraphMark>();
            return setting != null && (setting.Val?.Value ?? true);
        }
        set {
            Compatibility? compatibility = _wordprocessingDocument.MainDocumentPart?
                .DocumentSettingsPart?.Settings?.GetFirstChild<Compatibility>();
            if (compatibility == null) {
                compatibility = new Compatibility();
                GetSettings().AddChild(compatibility, true);
            }
            SplitPageBreakAndParagraphMark? setting = compatibility.GetFirstChild<SplitPageBreakAndParagraphMark>();
            if (setting == null) {
                setting = new SplitPageBreakAndParagraphMark();
                compatibility.AddChild(setting, true);
            }
            setting.Val = value;
        }
    }
}
