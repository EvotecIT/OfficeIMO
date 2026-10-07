using System.Linq;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class WordOdtEndnotePlacementLossTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void EndnotePlacementLossReflectsAuthoredConfiguration(bool authoredPosition, bool referencedNote) {
        using WordDocument source = WordDocument.Create();
        WordParagraph paragraph = source.AddParagraph("Anchor");
        if (referencedNote) paragraph.AddEndNote("Endnote body");
        if (authoredPosition) source.AddEndnoteProperties(position: WordEndnotePosition.SectionEnd);

        var conversion = source.ToOpenDocumentResult();
        bool reportsPlacement = conversion.Report.Mappings.Any(mapping => mapping.Feature == "note-numbering-placement");
        Assert.Equal(authoredPosition, reportsPlacement);
        Assert.Equal(authoredPosition ? WordEndnotePosition.SectionEnd : WordEndnotePosition.DocumentEnd,
            source.EndnoteSettings.Position);
    }
}
