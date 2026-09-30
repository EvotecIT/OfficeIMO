using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfLosslessEncodingRegressionTests {
    [Theory]
    [InlineData(0)]
    [InlineData(2)]
    public void Edits_Encode_Unicode_With_Local_Fallback_State_And_Restore_Surrounding_State(int count) {
        string fallback = new string('?', count);
        string input = @"{\rtf1\ansi\uc" + count + @" {\b Target}\u937" + fallback + @"\par}";
        RtfLosslessEditor editor = RtfDocument.Read(input).EditLossless();
        editor.ReplaceText("Target", "ż 😀");
        Assert.Equal("ż 😀Ω", editor.ToReadResult().Document.Paragraphs[0].ToPlainText());
        Assert.Contains(@"\u937" + fallback, editor.ToRtf(), StringComparison.Ordinal);
        editor.AppendParagraph("Zażółć 😀");
        Assert.Equal("Zażółć 😀", editor.ToReadResult().Document.Paragraphs.Last().ToPlainText());
        editor.InsertRootParagraph(editor.RootNodeCount, "Śródka");
        Assert.Equal("Śródka", editor.ToReadResult().Document.Paragraphs.Last().ToPlainText());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(2)]
    public void Unicode_Metadata_Resource_And_Variable_Edits_Use_Their_Encoded_Fallback_Count(int count) {
        RtfLosslessEditor editor = RtfDocument.Read(@"{\rtf1\ansi\uc" + count + @" Body\par}").EditLossless();
        editor.SetInfo(RtfDocumentInfoField.Title, "Zażółć");
        editor.SetFont(3, "Czcionka ż");
        editor.SetStyleName(5, "Śródka");
        editor.SetDocumentVariable("żName", "Łódź");
        editor.SetRevisionAuthor(0, "Przemysław");
        RtfDocument document = editor.ToReadResult().Document;
        Assert.Equal("Zażółć", document.Info.Title);
        Assert.Equal("Czcionka ż", document.Fonts.Single(font => font.Id == 3).Name);
        Assert.Equal("Śródka", document.Styles.Single(style => style.Id == 5).Name);
        Assert.Equal("żName", Assert.Single(document.DocumentVariables).Name);
        Assert.Equal("Łódź", document.DocumentVariables[0].Value);
        Assert.Equal("Przemysław", Assert.Single(document.RevisionAuthors).Name);
        Assert.Equal("Body", Assert.Single(document.Paragraphs).ToPlainText());
    }
}
