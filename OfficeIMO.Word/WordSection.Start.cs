using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordSection {
    /// <summary>
    /// Gets or sets how this section starts relative to the preceding section.
    /// An omitted section type uses <see cref="WordSectionBreakType.NextPage"/>.
    /// </summary>
    public WordSectionBreakType BreakType {
        get => (_sectionProperties.GetFirstChild<SectionType>()?.Val?.Value ?? SectionMarkValues.NextPage).ToOfficeEnum();
        set {
            SectionMarkValues sectionMark = value.ToOpenXml();
            _sectionProperties.RemoveAllChildren<SectionType>();
            _sectionProperties.AddChild(new SectionType { Val = sectionMark }, true);
        }
    }

    /// <summary>Resolves an authored numbering restart after Word's odd/even section-start constraint.</summary>
    internal int? GetEffectivePageNumberStart() {
        if (_sectionProperties.GetFirstChild<PageNumberType>()?.Start?.Value is not int start || start <= 0) return null;
        if (_document.Sections.IndexOf(this) > 0 &&
            ((BreakType == WordSectionBreakType.OddPage && start % 2 == 0) ||
             (BreakType == WordSectionBreakType.EvenPage && start % 2 != 0))) return checked(start + 1);
        return start;
    }
}
