using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("left")]
    [InlineData("right")]
    [InlineData("firstLine")]
    [InlineData("hanging")]
    public void ParagraphIndentation_ClearRemovesTheSavedAttribute(string attribute) {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Inherited indentation");
        switch (attribute) {
            case "left": paragraph.IndentationBefore = global::OfficeIMO.OfficeLength.Points(19.65D); paragraph.IndentationBefore= null; break;
            case "right": paragraph.IndentationAfter = global::OfficeIMO.OfficeLength.Points(19.65D); paragraph.IndentationAfter= null; break;
            case "firstLine": paragraph.IndentationFirstLine = global::OfficeIMO.OfficeLength.Points(19.65D); paragraph.IndentationFirstLine= null; break;
            case "hanging": paragraph.IndentationHanging = global::OfficeIMO.OfficeLength.Points(19.65D); paragraph.IndentationHanging= null; break;
        }
        using var stream = new MemoryStream(document.ToBytes());
        using var saved = WordprocessingDocument.Open(stream, false);
        Indentation indentation = saved.MainDocumentPart!.Document!.Body!
            .Elements<Paragraph>().Single().ParagraphProperties!.Indentation!;
        Assert.DoesNotContain(indentation.GetAttributes(), value => value.LocalName == attribute);
        Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(saved));
    }
}
