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
        editor.SetInfo(RtfDocumentInfoField.HyperlinkBase, "https://example.test/");

        const string expected = @"{\rtf1\ansi{\info{\title New \{title\} \u380?}{\company Evotec}{\hlinkbase https://example.test/}}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        Assert.Equal("New {title} ż", read.Document.Info.Title);
        Assert.Equal("Evotec", read.Document.Info.Company);
        Assert.Equal("https://example.test/", read.Document.Info.HyperlinkBase);
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
    public void SetInfo_Removes_Info_Group_When_Last_Field_Is_Removed() {
        const string rtf = @"{\rtf1\ansi{\info{\title Gone}}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetInfo(RtfDocumentInfoField.Title, null);

        Assert.Equal(@"{\rtf1\ansi\pard Body\par}", editor.ToRtf());
        Assert.Null(editor.ToReadResult().Document.Info.Title);
    }

    [Fact]
    public void SetInfoTimestamp_Replaces_Adds_And_Removes_Metadata_Without_Normalizing_Body() {
        const string rtf = @"{\rtf1\ansi{\info{\title Keep}{\creatim\yr2020\mo1\dy2\hr3\min4\sec5}\edmins10\nofpages2}\pard Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetInfoTimestamp(RtfDocumentInfoTimestampField.Created, new DateTime(2026, 6, 15, 10, 20, 30));
        editor.SetInfoTimestamp(RtfDocumentInfoTimestampField.Revised, new DateTime(2026, 6, 16, 11, 21, 31));
        editor.SetInfoNumber(RtfDocumentInfoNumberField.EditingMinutes, 42);
        editor.SetInfoNumber(RtfDocumentInfoNumberField.NumberOfPages, null);
        editor.SetInfoNumber(RtfDocumentInfoNumberField.NumberOfWords, 120);

        const string expected = @"{\rtf1\ansi{\info{\title Keep}{\creatim\yr2026\mo6\dy15\hr10\min20\sec30}\edmins42{\revtim\yr2026\mo6\dy16\hr11\min21\sec31}\nofwords120}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfDocumentInfo info = editor.ToReadResult().Document.Info;
        Assert.Equal(new DateTime(2026, 6, 15, 10, 20, 30), info.Created);
        Assert.Equal(new DateTime(2026, 6, 16, 11, 21, 31), info.Revised);
        Assert.Equal(42, info.EditingMinutes);
        Assert.Null(info.NumberOfPages);
        Assert.Equal(120, info.NumberOfWords);
        Assert.Contains(@"Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetInfoTimestamp_And_Number_Create_Info_Group_When_Missing() {
        const string rtf = @"{\rtf1\ansi\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetInfoTimestamp(RtfDocumentInfoTimestampField.Printed, new DateTime(2026, 1, 2, 3, 4, 5));
        editor.SetInfoNumber(RtfDocumentInfoNumberField.InternalVersion, 123);

        Assert.Equal(@"{\rtf1\ansi{\info{\printim\yr2026\mo1\dy2\hr3\min4\sec5}\vern123}\pard Body\par}", editor.ToRtf());
        RtfDocumentInfo info = editor.ToReadResult().Document.Info;
        Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5), info.Printed);
        Assert.Equal(123, info.InternalVersion);
    }

    [Fact]
    public void SetDocumentVariable_Replaces_Adds_And_Removes_Metadata_Without_Normalizing_Body() {
        const string rtf = @"{\rtf1\ansi{\info{\title Keep}}{\*\userprops{\propname Owner}\proptype30{\staticval Evotec}}{\*\docvar {Client}{Old}}{\*\docvar {Remove}{Gone}}\pard Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetDocumentVariable("Client", "New {client} ż");
        editor.SetDocumentVariable("Region", "EMEA");
        editor.SetDocumentVariable("Remove", null);

        const string expected = @"{\rtf1\ansi{\info{\title Keep}}{\*\userprops{\propname Owner}\proptype30{\staticval Evotec}}{\*\docvar {Client}{New \{client\} \u380?}}{\*\docvar {Region}{EMEA}}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        Assert.Collection(
            read.Document.DocumentVariables,
            variable => {
                Assert.Equal("Client", variable.Name);
                Assert.Equal("New {client} ż", variable.Value);
            },
            variable => {
                Assert.Equal("Region", variable.Name);
                Assert.Equal("EMEA", variable.Value);
            });
        Assert.Equal("Keep", read.Document.Info.Title);
        Assert.Contains(@"Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetDocumentVariable_Creates_DocVar_After_Header_When_Metadata_Is_Missing() {
        const string rtf = @"{\rtf1\ansi\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetDocumentVariable("Mode", string.Empty);

        Assert.Equal(@"{\rtf1\ansi{\*\docvar {Mode}{}}\pard Body\par}", editor.ToRtf());
        RtfDocumentVariable variable = Assert.Single(editor.ToReadResult().Document.DocumentVariables);
        Assert.Equal("Mode", variable.Name);
        Assert.Equal(string.Empty, variable.Value);
    }

    [Fact]
    public void SetUserProperty_Replaces_Adds_And_Removes_Metadata_Without_Normalizing_Body() {
        const string rtf = @"{\rtf1\ansi{\info{\title Keep}}{\*\userprops{\propname Client}\proptype30{\staticval Old}{\propname Remove}\proptype30{\staticval Gone}{\propname External}{\linkval Sheet1!A1}}{\*\docvar {Mode}{Draft}}\pard Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetUserProperty(RtfUserProperty.Text("Client", "New {client} ż"));
        editor.SetUserProperty(RtfUserProperty.Boolean("Approved", true));
        editor.RemoveUserProperty("Remove");

        const string expected = @"{\rtf1\ansi{\info{\title Keep}}{\*\userprops{\propname Client}\proptype30{\staticval New \{client\} \u380?}{\propname External}{\linkval Sheet1!A1}{\propname Approved}\proptype11{\staticval 1}}{\*\docvar {Mode}{Draft}}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        Assert.Collection(
            read.Document.UserProperties,
            property => {
                Assert.Equal("Client", property.Name);
                Assert.Equal(RtfUserProperty.TextType, property.TypeCode);
                Assert.Equal("New {client} ż", property.StaticValue);
            },
            property => {
                Assert.Equal("External", property.Name);
                Assert.Equal("Sheet1!A1", property.LinkedValue);
            },
            property => {
                Assert.Equal("Approved", property.Name);
                Assert.Equal(RtfUserProperty.BooleanType, property.TypeCode);
                Assert.Equal("1", property.StaticValue);
            });
        Assert.DoesNotContain(read.Document.UserProperties, property => property.Name == "Remove");
        Assert.Equal("Draft", Assert.Single(read.Document.DocumentVariables).Value);
        Assert.Contains(@"Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetUserProperty_Creates_UserProps_Before_Existing_DocVars() {
        const string rtf = @"{\rtf1\ansi{\*\docvar {Mode}{Draft}}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetUserProperty("Client", "Contoso");

        Assert.Equal(@"{\rtf1\ansi{\*\userprops{\propname Client}\proptype30{\staticval Contoso}}{\*\docvar {Mode}{Draft}}\pard Body\par}", editor.ToRtf());
        RtfUserProperty property = Assert.Single(editor.ToReadResult().Document.UserProperties);
        Assert.Equal("Client", property.Name);
        Assert.Equal("Contoso", property.StaticValue);
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
