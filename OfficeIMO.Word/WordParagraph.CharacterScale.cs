using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>
    /// Gets or sets the directly authored character width as a percentage between 1 and 600.
    /// A value of 100 restores normal width; null removes the direct value and restores style inheritance.
    /// Character height and character spacing are independent of this setting.
    /// </summary>
    public int? CharacterScale {
        get => checked((int?)ScopedRunProperties?.GetFirstChild<CharacterScale>()?.Val?.Value);
        set {
            if (value.HasValue && (value.Value < 1 || value.Value > 600))
                throw new ArgumentOutOfRangeException(nameof(value), "Character width must be between 1 and 600 percent.");
            RunProperties properties = IsHyperLink && _stdRun == null
                ? VerifyRunProperties(Hyperlink!._hyperlink!, Hyperlink._run!, Hyperlink._runProperties)
                : VerifyRunProperties();
            if (value.HasValue) properties.CharacterScale = new CharacterScale { Val = value.Value };
            else properties.CharacterScale?.Remove();
        }
    }
}
