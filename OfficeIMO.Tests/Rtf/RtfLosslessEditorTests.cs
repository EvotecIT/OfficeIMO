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
    public void SetGenerator_Replaces_Duplicates_And_Preserves_Body() {
        const string rtf = @"{\rtf1\ansi{\*\generator Old;}{\*\generator Duplicate;}{\info{\title Keep}}\pard Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetGenerator("New {generator} ż");

        const string expected = @"{\rtf1\ansi{\*\generator New \{generator\} \u380?;}{\info{\title Keep}}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());
        Assert.Equal("New {generator} ż", editor.ToReadResult().Document.Info.Generator);
        Assert.Equal("Keep", editor.ToReadResult().Document.Info.Title);
        Assert.Contains(@"Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetGenerator_Creates_And_Removes_Root_Generator_Group() {
        const string rtf = @"{\rtf1\ansi\deff0\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetGenerator("OfficeIMO");

        Assert.Equal(@"{\rtf1\ansi\deff0{\*\generator OfficeIMO;}\pard Body\par}", editor.ToRtf());
        Assert.Equal("OfficeIMO", editor.ToReadResult().Document.Info.Generator);

        editor.SetGenerator(null);

        Assert.Equal(rtf, editor.ToRtf());
        Assert.Null(editor.ToReadResult().Document.Info.Generator);
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
    public void SetXmlNamespace_Replaces_Adds_And_Removes_Metadata_Without_Normalizing_Body() {
        const string rtf = @"{\rtf1\ansi{\*\xmlnstbl{\xmlns1 old;}{\xmlns2 remove;}{\xmlns1 duplicate;}}{\info{\title Keep}}\pard Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetXmlNamespace(1, "urn:new {ns} ż");
        editor.SetXmlNamespace(3, "urn:add");
        editor.RemoveXmlNamespace(2);

        const string expected = @"{\rtf1\ansi{\*\xmlnstbl{\xmlns1 urn:new \{ns\} \u380?;}{\xmlns3 urn:add;}}{\info{\title Keep}}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        Assert.Collection(
            read.Document.XmlNamespaces,
            xmlNamespace => {
                Assert.Equal(1, xmlNamespace.Id);
                Assert.Equal("urn:new {ns} ż", xmlNamespace.Uri);
            },
            xmlNamespace => {
                Assert.Equal(3, xmlNamespace.Id);
                Assert.Equal("urn:add", xmlNamespace.Uri);
            });
        Assert.Equal("Keep", read.Document.Info.Title);
        Assert.Contains(@"Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetXmlNamespace_Creates_And_Removes_Namespace_Table() {
        const string rtf = @"{\rtf1\ansi{\fonttbl{\f0 Calibri;}}{\info{\title Keep}}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetXmlNamespace(7, "urn:created");

        Assert.Equal(@"{\rtf1\ansi{\fonttbl{\f0 Calibri;}}{\*\xmlnstbl{\xmlns7 urn:created;}}{\info{\title Keep}}\pard Body\par}", editor.ToRtf());
        RtfXmlNamespace xmlNamespace = Assert.Single(editor.ToReadResult().Document.XmlNamespaces);
        Assert.Equal(7, xmlNamespace.Id);
        Assert.Equal("urn:created", xmlNamespace.Uri);

        editor.RemoveXmlNamespace(7);

        Assert.Equal(rtf, editor.ToRtf());
        Assert.Empty(editor.ToReadResult().Document.XmlNamespaces);
    }

    [Fact]
    public void SetFileReference_Replaces_Adds_And_Removes_Metadata_Without_Normalizing_Body() {
        const string rtf = @"{\rtf1\ansi{\*\filetbl{\file\fid0\frelative18\fvalidntfs Old.docx}{\file\fid1\fvalidmac Remove.doc}{\file\fid0 Duplicate.docx}}{\*\xmlnstbl{\xmlns1 urn:keep;}}{\info{\title Keep}}\pard Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetFileReference(new RtfFileReference(0, @"C:\New {file} ż.docx") {
            RelativePathStart = 3,
            OperatingSystemNumber = 42,
            Sources = RtfFileSource.Ntfs | RtfFileSource.Network
        });
        editor.SetFileReference(new RtfFileReference(2, @"\\Server\Share\Added.docx") {
            Sources = RtfFileSource.Network
        });
        editor.RemoveFileReference(1);

        const string expected = @"{\rtf1\ansi{\*\filetbl{\file\fid0\frelative3\fosnum42\fvalidntfs\fnetwork C:\\New \{file\} \u380?.docx}{\file\fid2\fnetwork \\\\Server\\Share\\Added.docx}}{\*\xmlnstbl{\xmlns1 urn:keep;}}{\info{\title Keep}}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        Assert.Collection(
            read.Document.FileReferences,
            file => {
                Assert.Equal(0, file.Id);
                Assert.Equal(@"C:\New {file} ż.docx", file.Path);
                Assert.Equal(3, file.RelativePathStart);
                Assert.Equal(42, file.OperatingSystemNumber);
                Assert.Equal(RtfFileSource.Ntfs | RtfFileSource.Network, file.Sources);
            },
            file => {
                Assert.Equal(2, file.Id);
                Assert.Equal(@"\\Server\Share\Added.docx", file.Path);
                Assert.Null(file.RelativePathStart);
                Assert.Null(file.OperatingSystemNumber);
                Assert.Equal(RtfFileSource.Network, file.Sources);
            });
        Assert.Equal("urn:keep", Assert.Single(read.Document.XmlNamespaces).Uri);
        Assert.Equal("Keep", read.Document.Info.Title);
        Assert.Contains(@"Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetFileReference_Creates_And_Removes_File_Table() {
        const string rtf = @"{\rtf1\ansi{\fonttbl{\f0 Calibri;}}{\*\xmlnstbl{\xmlns1 urn:keep;}}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetFileReference(new RtfFileReference(4, "created.docx") {
            Sources = RtfFileSource.Dos
        });

        Assert.Equal(@"{\rtf1\ansi{\fonttbl{\f0 Calibri;}}{\*\filetbl{\file\fid4\fvaliddos created.docx}}{\*\xmlnstbl{\xmlns1 urn:keep;}}\pard Body\par}", editor.ToRtf());
        RtfFileReference file = Assert.Single(editor.ToReadResult().Document.FileReferences);
        Assert.Equal(4, file.Id);
        Assert.Equal("created.docx", file.Path);
        Assert.Equal(RtfFileSource.Dos, file.Sources);

        editor.RemoveFileReference(4);

        Assert.Equal(rtf, editor.ToRtf());
        Assert.Empty(editor.ToReadResult().Document.FileReferences);
    }

    [Fact]
    public void RevisionSaveIds_Update_Root_Add_Remove_And_Preserve_Body() {
        const string rtf = @"{\rtf1\ansi{\*\rsidtbl\rsidroot7\rsid15\rsid1024\rsid15}{\info{\title Keep}}\pard\pararsid20 Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetRevisionRootSaveId(9);
        editor.AddRevisionSaveId(2048);
        editor.AddRevisionSaveId(15);
        editor.RemoveRevisionSaveId(1024);

        const string expected = @"{\rtf1\ansi{\*\rsidtbl\rsidroot9\rsid15\rsid2048}{\info{\title Keep}}\pard\pararsid20 Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        Assert.Equal(9, read.Document.RevisionRootSaveId);
        Assert.Equal(new[] { 15, 2048 }, read.Document.RevisionSaveIds);
        Assert.Equal("Keep", read.Document.Info.Title);
        Assert.Contains(@"\pararsid20 Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void RevisionSaveIds_Create_And_Remove_Rsid_Table() {
        const string rtf = @"{\rtf1\ansi{\*\revtbl{Alice;}}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetRevisionRootSaveId(7);
        editor.AddRevisionSaveId(15);

        Assert.Equal(@"{\rtf1\ansi{\*\revtbl{Alice;}}{\*\rsidtbl\rsidroot7\rsid15}\pard Body\par}", editor.ToRtf());
        RtfReadResult read = editor.ToReadResult();
        Assert.Equal(7, read.Document.RevisionRootSaveId);
        Assert.Equal(new[] { 15 }, read.Document.RevisionSaveIds);
        Assert.Equal("Alice", Assert.Single(read.Document.RevisionAuthors).Name);

        editor.SetRevisionRootSaveId(null);
        editor.RemoveRevisionSaveId(15);

        Assert.Equal(rtf, editor.ToRtf());
        RtfReadResult removed = editor.ToReadResult();
        Assert.Null(removed.Document.RevisionRootSaveId);
        Assert.Empty(removed.Document.RevisionSaveIds);
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
