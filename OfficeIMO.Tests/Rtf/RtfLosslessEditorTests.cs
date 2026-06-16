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
    public void SetRootSettings_Replaces_Duplicates_And_Preserves_Tables_And_Body() {
        const string rtf = @"{\rtf1\ansi\mac\ansicpg1250\deff1\uc2{\fonttbl{\f0 Calibri;}{\f1 Consolas;}}{\info{\title Keep}}\pard Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetCharacterSet(RtfDocumentCharacterSet.Mac, 1252);
        editor.SetDefaultFont(0);
        editor.SetUnicodeSkipCount(0);

        const string expected = @"{\rtf1\mac\ansicpg1252\deff0\uc0{\fonttbl{\f0 Calibri;}{\f1 Consolas;}}{\info{\title Keep}}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfDocumentSettings settings = editor.ToReadResult().Document.Settings;
        Assert.Equal(RtfDocumentCharacterSet.Mac, settings.CharacterSet);
        Assert.Equal(1252, settings.AnsiCodePage);
        Assert.Equal(0, settings.DefaultFontId);
        Assert.Equal(0, settings.UnicodeSkipCount);
        Assert.Contains(@"{\fonttbl{\f0 Calibri;}{\f1 Consolas;}}", editor.ToRtf(), StringComparison.Ordinal);
        Assert.Contains(@"Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetRootSettings_Creates_And_Removes_Optional_Header_Controls_Before_Metadata() {
        const string rtf = @"{\rtf1{\info{\title Keep}}\pard Body \u380??\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetCharacterSet(RtfDocumentCharacterSet.Ansi, 1250);
        editor.SetDefaultFont(2);
        editor.SetUnicodeSkipCount(2);

        Assert.Equal(@"{\rtf1\ansi\ansicpg1250\deff2\uc2{\info{\title Keep}}\pard Body \u380??\par}", editor.ToRtf());
        RtfDocumentSettings settings = editor.ToReadResult().Document.Settings;
        Assert.Equal(RtfDocumentCharacterSet.Ansi, settings.CharacterSet);
        Assert.Equal(1250, settings.AnsiCodePage);
        Assert.Equal(2, settings.DefaultFontId);
        Assert.Equal(2, settings.UnicodeSkipCount);

        editor.SetAnsiCodePage(null);
        editor.SetDefaultFont(null);
        editor.SetUnicodeSkipCount(null);

        Assert.Equal(@"{\rtf1\ansi{\info{\title Keep}}\pard Body \u380??\par}", editor.ToRtf());
        settings = editor.ToReadResult().Document.Settings;
        Assert.Equal(RtfDocumentCharacterSet.Ansi, settings.CharacterSet);
        Assert.Null(settings.AnsiCodePage);
        Assert.Null(settings.DefaultFontId);
        Assert.Null(settings.UnicodeSkipCount);
    }

    [Fact]
    public void SetRootPageSetup_Replaces_Duplicates_And_Preserves_Metadata_And_Body() {
        const string rtf = @"{\rtf1\ansi\paperw1000\paperw1111\paperh2000\psz4\binfsxn2\binsxn3\margl100\margr200\margt300\margb400\headery50\footery60\gutter70\rtlgutter\landscape\titlepg{\fonttbl{\f0 Calibri;}}{\info{\title Keep}}\pard Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetPageSize(12240, 15840);
        editor.SetPrinterPaper(9, 7, 8);
        editor.SetMargins(720, 1440, 1080, 900);
        editor.SetGutterWidth(180);
        editor.SetHeaderFooterDistance(360, 540);
        editor.SetRtlGutter(false);
        editor.SetLandscape(false);
        editor.SetDifferentFirstPageHeaderFooter(false);

        const string expected = @"{\rtf1\ansi\paperw12240\paperh15840\psz9\binfsxn7\binsxn8\margl720\margr1440\margt1080\margb900\gutter180\headery360\footery540{\fonttbl{\f0 Calibri;}}{\info{\title Keep}}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        RtfPageSetup pageSetup = read.Document.PageSetup;
        Assert.Equal(12240, pageSetup.PaperWidthTwips);
        Assert.Equal(15840, pageSetup.PaperHeightTwips);
        Assert.Equal(9, pageSetup.PrinterPaperSize);
        Assert.Equal(7, pageSetup.FirstPagePaperSource);
        Assert.Equal(8, pageSetup.OtherPagesPaperSource);
        Assert.Equal(720, pageSetup.MarginLeftTwips);
        Assert.Equal(1440, pageSetup.MarginRightTwips);
        Assert.Equal(1080, pageSetup.MarginTopTwips);
        Assert.Equal(900, pageSetup.MarginBottomTwips);
        Assert.Equal(180, pageSetup.GutterWidthTwips);
        Assert.Equal(360, pageSetup.HeaderDistanceTwips);
        Assert.Equal(540, pageSetup.FooterDistanceTwips);
        Assert.False(pageSetup.RtlGutter);
        Assert.False(pageSetup.Landscape);
        Assert.False(pageSetup.DifferentFirstPageHeaderFooter);
        Assert.Equal("Keep", read.Document.Info.Title);
        Assert.Contains(@"Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetRootPageSetup_Creates_And_Removes_Optional_Controls_Before_Metadata() {
        const string rtf = @"{\rtf1\ansi{\info{\title Keep}}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetPageSize(8000, 10000);
        editor.SetMargins(leftTwips: 720, topTwips: 360);
        editor.SetHeaderFooterDistance(300, 420);
        editor.SetRtlGutter();
        editor.SetLandscape();
        editor.SetDifferentFirstPageHeaderFooter();

        const string expected = @"{\rtf1\ansi\paperw8000\paperh10000\margl720\margt360\headery300\footery420\rtlgutter\landscape\titlepg{\info{\title Keep}}\pard Body\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfPageSetup pageSetup = editor.ToReadResult().Document.PageSetup;
        Assert.Equal(8000, pageSetup.PaperWidthTwips);
        Assert.Equal(10000, pageSetup.PaperHeightTwips);
        Assert.Equal(720, pageSetup.MarginLeftTwips);
        Assert.Null(pageSetup.MarginRightTwips);
        Assert.Equal(360, pageSetup.MarginTopTwips);
        Assert.Null(pageSetup.MarginBottomTwips);
        Assert.True(pageSetup.RtlGutter);
        Assert.True(pageSetup.Landscape);
        Assert.True(pageSetup.DifferentFirstPageHeaderFooter);

        editor.SetPageSize(null, null);
        editor.SetMargins();
        editor.SetHeaderFooterDistance();
        editor.SetRtlGutter(false);
        editor.SetLandscape(false);
        editor.SetDifferentFirstPageHeaderFooter(false);

        Assert.Equal(rtf, editor.ToRtf());
        pageSetup = editor.ToReadResult().Document.PageSetup;
        Assert.Null(pageSetup.PaperWidthTwips);
        Assert.Null(pageSetup.PaperHeightTwips);
        Assert.Null(pageSetup.MarginLeftTwips);
        Assert.Null(pageSetup.MarginTopTwips);
        Assert.Null(pageSetup.HeaderDistanceTwips);
        Assert.Null(pageSetup.FooterDistanceTwips);
        Assert.False(pageSetup.RtlGutter);
        Assert.False(pageSetup.Landscape);
        Assert.False(pageSetup.DifferentFirstPageHeaderFooter);
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
    public void SetFont_Replaces_Adds_And_Removes_Font_Table_Entries_Without_Normalizing_Body() {
        const string rtf = @"{\rtf1\ansi{\fonttbl{\f0\fswiss Calibri;}{\f1\froman Old;}{\f1 Duplicate;}}{\info{\title Keep}}\pard\f1 Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetFont(new RtfFont(1, "New {Font} ż") {
            Family = RtfFontFamily.Modern,
            Charset = 238,
            Pitch = 1,
            CodePage = 1250,
            Bias = 0,
            Panose = "020F0502020204030204",
            NonTaggedName = "New Font",
            AlternateName = "Fallback",
            Embedding = new RtfFontEmbedding {
                Type = RtfEmbeddedFontType.TrueType,
                FileCodePage = 1250,
                FileName = "NewFont.ttf",
                Data = new byte[] { 0x01, 0x02, 0xFF }
            }
        });
        editor.SetFont(2, "Added");
        editor.RemoveFont(0);

        const string expected = @"{\rtf1\ansi{\fonttbl{\f1\fmodern\fcharset238\fprq1\cpg1250\fbias0{\*\panose 020F0502020204030204}{\*\fname New Font}{\*\fontemb\fttruetype{\*\fontfile\cpg1250 NewFont.ttf} 0102ff} New \{Font\} \u380?{\*\falt Fallback};}{\f2 Added;}}{\info{\title Keep}}\pard\f1 Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        Assert.Collection(
            read.Document.Fonts,
            font => {
                Assert.Equal(1, font.Id);
                Assert.Equal("New {Font} ż", font.Name);
                Assert.Equal(RtfFontFamily.Modern, font.Family);
                Assert.Equal(238, font.Charset);
                Assert.Equal(1, font.Pitch);
                Assert.Equal(1250, font.CodePage);
                Assert.Equal(0, font.Bias);
                Assert.Equal("020F0502020204030204", font.Panose);
                Assert.Equal("New Font", font.NonTaggedName);
                Assert.Equal("Fallback", font.AlternateName);
                Assert.NotNull(font.Embedding);
                Assert.Equal(RtfEmbeddedFontType.TrueType, font.Embedding!.Type);
                Assert.Equal(1250, font.Embedding.FileCodePage);
                Assert.Equal("NewFont.ttf", font.Embedding.FileName);
                Assert.Equal(new byte[] { 0x01, 0x02, 0xFF }, font.Embedding.Data);
            },
            font => {
                Assert.Equal(2, font.Id);
                Assert.Equal("Added", font.Name);
            });
        Assert.Equal("Keep", read.Document.Info.Title);
        Assert.Contains(@"\pard\f1 Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetFont_Creates_And_Removes_Font_Table_Before_Metadata() {
        const string rtf = @"{\rtf1\ansi{\info{\title Keep}}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetFont(new RtfFont(0, "Calibri") {
            Family = RtfFontFamily.Swiss,
            Charset = 0
        });

        Assert.Equal(@"{\rtf1\ansi{\fonttbl{\f0\fswiss\fcharset0 Calibri;}}{\info{\title Keep}}\pard Body\par}", editor.ToRtf());
        RtfFont font = Assert.Single(editor.ToReadResult().Document.Fonts);
        Assert.Equal(0, font.Id);
        Assert.Equal("Calibri", font.Name);
        Assert.Equal(RtfFontFamily.Swiss, font.Family);
        Assert.Equal(0, font.Charset);

        editor.RemoveFont(0);

        Assert.Equal(rtf, editor.ToRtf());
        Assert.Equal("Calibri", Assert.Single(editor.ToReadResult().Document.Fonts).Name);
    }

    [Fact]
    public void SetStyleName_Renames_Adds_Deduplicates_And_Removes_Styles_Without_Normalizing_Body() {
        const string rtf = @"{\rtf1\ansi{\fonttbl{\f0 Calibri;}}{\stylesheet{\s1\sbasedon0\snext1\b Old;}{\*\cs2\additive Link;}{\s1 Duplicate;}}\pard\s1 Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetStyleName(1, "Heading {One} ż");
        editor.SetStyleName(3, "Added", RtfStyleKind.Character);
        editor.RemoveStyle(2, RtfStyleKind.Character);

        const string expected = @"{\rtf1\ansi{\fonttbl{\f0 Calibri;}}{\stylesheet{\s1\sbasedon0\snext1\b Heading \{One\} \u380?;}{\*\cs3 Added;}}\pard\s1 Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        Assert.Collection(
            read.Document.Styles,
            style => {
                Assert.Equal(1, style.Id);
                Assert.Equal(RtfStyleKind.Paragraph, style.Kind);
                Assert.Equal("Heading {One} ż", style.Name);
                Assert.Equal(0, style.BasedOnStyleId);
                Assert.Equal(1, style.NextStyleId);
                Assert.True(style.Bold);
            },
            style => {
                Assert.Equal(3, style.Id);
                Assert.Equal(RtfStyleKind.Character, style.Kind);
                Assert.Equal("Added", style.Name);
            });
        Assert.Contains(@"\pard\s1 Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetStyleName_Creates_And_Removes_Stylesheet_Before_Revision_Table() {
        const string rtf = @"{\rtf1\ansi{\*\revtbl{Alice;}}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetStyleName(9, "Table Grid", RtfStyleKind.Table);

        Assert.Equal(@"{\rtf1\ansi{\stylesheet{\*\ts9 Table Grid;}}{\*\revtbl{Alice;}}\pard Body\par}", editor.ToRtf());
        RtfStyle style = Assert.Single(editor.ToReadResult().Document.Styles);
        Assert.Equal(9, style.Id);
        Assert.Equal(RtfStyleKind.Table, style.Kind);
        Assert.Equal("Table Grid", style.Name);

        editor.RemoveStyle(9, RtfStyleKind.Table);

        Assert.Equal(rtf, editor.ToRtf());
        Assert.Empty(editor.ToReadResult().Document.Styles);
    }

    [Fact]
    public void SetRevisionAuthor_Replaces_Adds_And_Removes_Authors_Without_Normalizing_Body() {
        const string rtf = @"{\rtf1\ansi{\*\revtbl{Alice;}{Bob;}{Remove;}}{\*\rsidtbl\rsidroot7}\pard\revised\revauth1 Body \'80\revised0\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetRevisionAuthor(1, "Bób {Two}");
        editor.SetRevisionAuthor(3, "Dana");
        editor.RemoveRevisionAuthor(2);

        const string expected = @"{\rtf1\ansi{\*\revtbl{Alice;}{B\u243?b \{Two\};}{Dana;}}{\*\rsidtbl\rsidroot7}\pard\revised\revauth1 Body \'80\revised0\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        Assert.Collection(
            read.Document.RevisionAuthors,
            author => Assert.Equal("Alice", author.Name),
            author => Assert.Equal("Bób {Two}", author.Name),
            author => Assert.Equal("Dana", author.Name));
        Assert.Equal(7, read.Document.RevisionRootSaveId);
        Assert.Contains(@"\pard\revised\revauth1 Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetRevisionAuthor_Creates_And_Removes_Revision_Table_Before_SaveIds() {
        const string rtf = @"{\rtf1\ansi{\*\rsidtbl\rsidroot7}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetRevisionAuthor(0, "Alice");
        editor.SetRevisionAuthor(1, "Bob");

        Assert.Equal(@"{\rtf1\ansi{\*\revtbl{Alice;}{Bob;}}{\*\rsidtbl\rsidroot7}\pard Body\par}", editor.ToRtf());
        Assert.Collection(
            editor.ToReadResult().Document.RevisionAuthors,
            author => Assert.Equal("Alice", author.Name),
            author => Assert.Equal("Bob", author.Name));

        editor.RemoveRevisionAuthor(1);
        editor.RemoveRevisionAuthor(0);

        Assert.Equal(rtf, editor.ToRtf());
        Assert.Empty(editor.ToReadResult().Document.RevisionAuthors);
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
    public void SetColor_Replaces_Adds_And_Removes_Color_Table_Entries_Without_Normalizing_Body() {
        const string rtf = @"{\rtf1\ansi{\fonttbl{\f0 Calibri;}}{\colortbl;\red1\green2\blue3;\red4\green5\blue6\caccenttwo\cshade25;}{\stylesheet{\s1\cf1 Style;}}\pard\cf1 Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetColor(1, new RtfColor(10, 20, 30) {
            ThemeColor = RtfThemeColor.AccentOne,
            Tint = 40
        });
        editor.SetColor(3, 7, 8, 9);
        editor.RemoveColor(2);

        const string expected = @"{\rtf1\ansi{\fonttbl{\f0 Calibri;}}{\colortbl;\red10\green20\blue30\caccentone\ctint40;\red7\green8\blue9;}{\stylesheet{\s1\cf1 Style;}}\pard\cf1 Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        Assert.Collection(
            read.Document.Colors,
            color => {
                Assert.Equal(10, color.Red);
                Assert.Equal(20, color.Green);
                Assert.Equal(30, color.Blue);
                Assert.Equal(RtfThemeColor.AccentOne, color.ThemeColor);
                Assert.Equal(40, color.Tint);
                Assert.Null(color.Shade);
            },
            color => {
                Assert.Equal(7, color.Red);
                Assert.Equal(8, color.Green);
                Assert.Equal(9, color.Blue);
                Assert.Null(color.ThemeColor);
            });
        Assert.Contains(@"\pard\cf1 Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetColor_Creates_And_Removes_Color_Table_Before_Stylesheet() {
        const string rtf = @"{\rtf1\ansi{\fonttbl{\f0 Calibri;}}{\stylesheet{\s1 Style;}}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetColor(1, new RtfColor(68, 114, 196) {
            ThemeColor = RtfThemeColor.Hyperlink,
            Shade = 25
        });

        Assert.Equal(@"{\rtf1\ansi{\fonttbl{\f0 Calibri;}}{\colortbl;\red68\green114\blue196\chyperlink\cshade25;}{\stylesheet{\s1 Style;}}\pard Body\par}", editor.ToRtf());
        RtfColor color = Assert.Single(editor.ToReadResult().Document.Colors);
        Assert.Equal(68, color.Red);
        Assert.Equal(114, color.Green);
        Assert.Equal(196, color.Blue);
        Assert.Equal(RtfThemeColor.Hyperlink, color.ThemeColor);
        Assert.Equal(25, color.Shade);

        editor.RemoveColor(1);

        Assert.Equal(rtf, editor.ToRtf());
        Assert.Empty(editor.ToReadResult().Document.Colors);
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

    [Fact]
    public void SaveLossless_Writes_Edited_Rtf_To_Stream_And_File_Without_Normalizing_Untouched_Bytes() {
        byte[] sourceBytes = new byte[] {
            123, 92, 114, 116, 102, 49, 92, 97, 110, 115, 105, 123, 92, 42, 92, 117, 110, 107, 110, 111, 119, 110, 32, 0x80, 125,
            92, 112, 97, 114, 100, 32, 79, 108, 100, 92, 112, 97, 114, 125
        };
        string outputPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".rtf");

        try {
            using var input = new MemoryStream(sourceBytes);
            RtfLosslessEditor editor = RtfDocument.Load(input).EditLossless();
            editor.ReplaceText("Old", "New");

            using var output = new MemoryStream();
            editor.SaveLossless(output);
            editor.SaveLossless(outputPath);

            byte[] expected = new byte[] {
                123, 92, 114, 116, 102, 49, 92, 97, 110, 115, 105, 123, 92, 42, 92, 117, 110, 107, 110, 111, 119, 110, 32, 0x80, 125,
                92, 112, 97, 114, 100, 32, 78, 101, 119, 92, 112, 97, 114, 125
            };
            Assert.Equal(expected, editor.ToBytesLossless());
            Assert.Equal(expected, output.ToArray());
            Assert.Equal(expected, File.ReadAllBytes(outputPath));
        } finally {
            if (File.Exists(outputPath)) {
                File.Delete(outputPath);
            }
        }
    }
}
