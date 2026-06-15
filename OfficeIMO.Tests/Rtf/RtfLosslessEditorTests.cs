using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfLosslessEditorTests {
    [Fact]
    public void ReplaceText_Edits_Visible_Text_And_Preserves_Structural_Destinations() {
        const string rtf = @"{\rtf1\ansi{\*\unknown Target}{\fonttbl{\f0 Target;}}{\*\userprops{\propname Target}\proptype30{\staticval Target}}{\*\docvar {Target}{Target}}\pard Target {\b Target}\par{\field{\*\fldinst HYPERLINK ""https://example.test/Target""}{\fldrslt Target}}}";
        const string expected = @"{\rtf1\ansi{\*\unknown Target}{\fonttbl{\f0 Target;}}{\*\userprops{\propname Target}\proptype30{\staticval Target}}{\*\docvar {Target}{Target}}\pard Replaced {\b Replaced}\par{\field{\*\fldinst HYPERLINK ""https://example.test/Target""}{\fldrslt Replaced}}}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        int replacements = editor.ReplaceText("Target", "Replaced");

        Assert.Equal(3, replacements);
        Assert.Equal(expected, editor.ToRtf());
        Assert.Contains("https://example.test/Target", editor.ToRtf(), StringComparison.Ordinal);
        Assert.Contains(@"{\*\userprops{\propname Target}\proptype30{\staticval Target}}", editor.ToRtf(), StringComparison.Ordinal);
        Assert.Contains(@"{\*\docvar {Target}{Target}}", editor.ToRtf(), StringComparison.Ordinal);
        IReadOnlyList<RtfParagraph> paragraphs = editor.ToReadResult().Document.Paragraphs;
        Assert.Collection(
            paragraphs,
            paragraph => Assert.Equal("Replaced Replaced", paragraph.ToPlainText()),
            paragraph => Assert.Equal("Replaced", paragraph.ToPlainText()));
    }

    [Fact]
    public void ReplaceText_Escapes_Inserted_Rtf_Text() {
        const string rtf = @"{\rtf1\ansi\pard Target\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.ReplaceText("Target", @"A{\B} ż");

        Assert.Equal(@"{\rtf1\ansi\pard A\{\\B\} \u380?\par}", editor.ToRtf());
        Assert.Equal(@"A{\B} ż", editor.ToReadResult().Document.Paragraphs[0].ToPlainText());
    }

    [Fact]
    public void SetInfo_Replaces_Adds_And_Removes_Metadata_Without_Normalizing_Body() {
        const string rtf = @"{\rtf1\ansi{\info{\title Old}{\author Someone}}\pard Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetInfo(RtfDocumentInfoField.Title, "New {title} ż");
        editor.SetInfo(RtfDocumentInfoField.Author, null);
        editor.SetInfo(RtfDocumentInfoField.Company, "Evotec");

        const string expected = @"{\rtf1\ansi{\info{\title New \{title\} \u380?}{\company Evotec}}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        Assert.Equal("New {title} ż", read.Document.Info.Title);
        Assert.Equal("Evotec", read.Document.Info.Company);
        Assert.Null(read.Document.Info.Author);
        Assert.Contains(@"Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetInfo_Creates_Info_Group_When_Missing() {
        const string rtf = @"{\rtf1\ansi\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetInfo(RtfDocumentInfoField.Title, "Created");

        Assert.Equal(@"{\rtf1\ansi{\info{\title Created}}\pard Body\par}", editor.ToRtf());
        Assert.Equal("Created", editor.ToReadResult().Document.Info.Title);
    }

    [Fact]
    public void AppendParagraph_Preserves_Existing_Syntax_And_Escapes_Text() {
        const string rtf = @"{\rtf1\ansi{\*\unknown Keep}{\pict\pngblip\bin3 abc}\pard Existing\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.AppendParagraph(@"Next {\line} ż");

        const string expected = @"{\rtf1\ansi{\*\unknown Keep}{\pict\pngblip\bin3 abc}\pard Existing\par\pard Next \{\\line\} \u380?\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        Assert.Collection(
            read.Document.Paragraphs,
            paragraph => Assert.Equal("Existing", paragraph.ToPlainText()),
            paragraph => Assert.Equal(@"Next {\line} ż", paragraph.ToPlainText()));
        Assert.IsType<RtfImage>(read.Document.Blocks[0]);
    }
}
