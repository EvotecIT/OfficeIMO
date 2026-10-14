using OfficeIMO.Word;
using Xunit;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void NumberedRunningStoryAdmissionPreservesMixedContentInOtherActiveVariants(bool footer, bool even) {
        using WordDocument document = CreateJoinedParagraphDocument();
        WordList list = document.AddList(WordListStyle.Numbered);
        WordParagraph template = list.AddItem("TEMPLATE");
        WordSection section = document.Sections[0];
        section.DifferentFirstPage = !even;
        section.DifferentOddAndEvenPages = even;
        WordHeaderFooter normal = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordHeaderFooter mixed = footer
            ? even ? document.FooterEvenOrCreate : document.FooterFirstOrCreate
            : even ? document.HeaderEvenOrCreate : document.HeaderFirstOrCreate;
        normal.AddParagraph("PLAIN-LIST")._paragraph.ParagraphProperties =
            (W.ParagraphProperties)template._paragraph.ParagraphProperties!.CloneNode(true);
        WordParagraph outer = mixed.AddParagraph("PREFIX");
        WordParagraph box = outer.AddTextBox("BOX", WordImageTextWrapping.Square).Paragraphs.Single();
        box._paragraph.ParagraphProperties = (W.ParagraphProperties)template._paragraph.ParagraphProperties!.CloneNode(true);
        box._paragraph.Descendants<W.Break>().ToList().ForEach(item => item.Remove());
        outer._paragraph.Append(new W.SimpleField(new W.Run(new W.Text("999"))) { Instruction = " PAGE " });
        outer.AddText("SUFFIX");
        template._paragraph.Remove();
        document.AddParagraph("BODY"); document.AddParagraph().AddBreak(WordBreakType.Page); document.AddParagraph("BODY");
        Assert.Empty(document.ValidateDocument());
        using var pdf = OpenJoinedParagraphPdf(document);
        Assert.Equal(2, pdf.NumberOfPages);
        string text = pdf.GetPage(even ? 2 : 1).Text;
        Assert.Contains("PREFIX", text); Assert.Contains("BOX", text); Assert.Contains("SUFFIX", text);
        Assert.DoesNotContain("999", text);
        Assert.Contains("PLAIN-LIST", pdf.GetPage(even ? 1 : 2).Text);
    }
}
