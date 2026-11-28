using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;
using OfficeIMO.Word.OpenXml;

namespace OfficeIMO.Word;

public partial class WordSection {
    /// <summary>
    /// Gets or sets how this section starts relative to the preceding section.
    /// An omitted section type uses <see cref="WordSectionBreakType.NextPage"/>.
    /// </summary>
    public WordSectionBreakType BreakType {
        get => WordSectionBreakTypeExtensions.FromOpenXmlToken(FindSectionChild(SectionNode, "type").Attribute(SectionNamespace, "val"));
        set => SetBreakType(value.ToOpenXmlToken());
    }

    private void SetBreakType(string token) {
        EnsureSectionChild("type").SetAttribute(SectionNamespace, "val", token);
    }

    /// <summary>Resolves an authored numbering restart after Word's odd/even section-start constraint.</summary>
    internal int? GetEffectivePageNumberStart() {
        if (OfficeXmlViewValues.ParseInt32(FindSectionChild(SectionNode, "pgNumType").Attribute(SectionNamespace, "start")) is not int start || start <= 0) return null;
        if (_document.Sections.IndexOf(this) > 0 &&
            ((BreakType == WordSectionBreakType.OddPage && start % 2 == 0) ||
             (BreakType == WordSectionBreakType.EvenPage && start % 2 != 0))) return checked(start + 1);
        return start;
    }
}
