using static OfficeIMO.Word.OpenXml.WordXmlEdits;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>
    /// Gets or sets the directly authored character width as a percentage between 1 and 600.
    /// A value of 100 restores normal width; null removes the direct value and restores style inheritance.
    /// Character height and character spacing are independent of this setting.
    /// </summary>
    public int? CharacterScale {
        get => int.TryParse(ChildValue(RunPropertiesNode, "w"), System.Globalization.NumberStyles.Integer,
            System.Globalization.CultureInfo.InvariantCulture, out int scale) ? scale : null;
        set {
            if (value.HasValue && (value.Value < 1 || value.Value > 600))
                throw new ArgumentOutOfRangeException(nameof(value), "Character width must be between 1 and 600 percent.");
            SetRunValue("w", value?.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
    }
}
