using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentReadWriteTests {
    [Fact]
    public void Read_Binds_Embedded_Object_Data_Result_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\pard Before {\object\objemb\objw100\objh200\objscalex75\objscaley80{\*\objclass Package}{\*\objname Attachment}{\*\objdata 010203ff}{\result Display}} after\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "RTF101");
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Before Display after", paragraph.ToPlainText());
        RtfObject rtfObject = Assert.IsType<RtfObject>(paragraph.Inlines[1]);
        Assert.Equal(RtfObjectKind.Embedded, rtfObject.Kind);
        Assert.Equal("Package", rtfObject.ClassName);
        Assert.Equal("Attachment", rtfObject.Name);
        Assert.Equal(new byte[] { 1, 2, 3, 255 }, rtfObject.Data);
        Assert.Equal(100, rtfObject.Width);
        Assert.Equal(200, rtfObject.Height);
        Assert.Equal(75, rtfObject.ScaleX);
        Assert.Equal(80, rtfObject.ScaleY);
        Assert.Equal("Display", rtfObject.Result.ToPlainText());
    }

    [Fact]
    public void Read_Binds_Object_Result_Image_And_Text_Together() {
        const string rtf = @"{\rtf1\ansi\pard Before {\object\objemb{\*\objdata 0102}{\result{\pict\pngblip\bin3 abc}Caption}} after\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Before Caption after", paragraph.ToPlainText());
        RtfObject rtfObject = Assert.IsType<RtfObject>(paragraph.Inlines[1]);
        Assert.NotNull(rtfObject.ResultImage);
        Assert.Equal(RtfImageFormat.Png, rtfObject.ResultImage!.Format);
        Assert.Equal(new byte[] { (byte)'a', (byte)'b', (byte)'c' }, rtfObject.ResultImage.Data);
        Assert.Equal("Caption", rtfObject.Result.ToPlainText());
    }

    [Fact]
    public void Write_Emits_Object_Result_Image_And_Text_Together() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Before ");
        RtfObject rtfObject = paragraph.AddObject(RtfObjectKind.Embedded, new byte[] { 1, 2 });
        rtfObject.ResultImage = new RtfImage(RtfImageFormat.Png, new byte[] { 0x01, 0x02, 0x03 });
        rtfObject.Result.AddText("Caption").SetBold();
        paragraph.AddText(" after");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"{\result {\pict\pngblip", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\b Caption\b0 ", rtf, StringComparison.Ordinal);
        RtfObject readObject = Assert.IsType<RtfObject>(Assert.Single(result.Document.Paragraphs).Inlines[1]);
        Assert.NotNull(readObject.ResultImage);
        Assert.Equal(new byte[] { 0x01, 0x02, 0x03 }, readObject.ResultImage!.Data);
        Assert.Equal("Caption", readObject.Result.ToPlainText());
        Assert.Contains(readObject.Result.Runs, run => run.Text == "Caption" && run.Bold);
    }

    [Fact]
    public void Read_Binds_Hidden_Text_Runs_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\pard Visible \v Hidden\v0  shown\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Visible Hidden shown", paragraph.ToPlainText());
        Assert.Contains(paragraph.Runs, run => run.Text == "Hidden" && run.Hidden);
        Assert.Contains(paragraph.Runs, run => run.Text.Contains("shown", StringComparison.Ordinal) && !run.Hidden);
    }

    [Fact]
    public void Write_And_Read_Double_Strike_And_Caps_Effects_Without_Leaking() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Double").SetDoubleStrike();
        paragraph.AddText(" Caps").SetCapsStyle(RtfCapsStyle.Caps);
        paragraph.AddText(" Small").SetCapsStyle(RtfCapsStyle.SmallCaps);
        paragraph.AddText(" plain");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"\striked Double", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\striked0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\caps  Caps", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\caps0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\scaps  Small", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\scaps0", rtf, StringComparison.Ordinal);

        RtfParagraph readParagraph = Assert.Single(result.Document.Paragraphs);
        Assert.Contains(readParagraph.Runs, run => run.Text == "Double" && run.DoubleStrike);
        Assert.Contains(readParagraph.Runs, run => run.Text == " Caps" && run.CapsStyle == RtfCapsStyle.Caps);
        Assert.Contains(readParagraph.Runs, run => run.Text == " Small" && run.CapsStyle == RtfCapsStyle.SmallCaps);
        RtfRun plain = readParagraph.Runs.Single(run => run.Text == " plain");
        Assert.False(plain.DoubleStrike);
        Assert.Equal(RtfCapsStyle.None, plain.CapsStyle);
    }

    [Fact]
    public void Write_And_Read_Outline_Shadow_Emboss_And_Imprint_Effects_Without_Leaking() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("Outline").SetOutline();
        paragraph.AddText(" Shadow").SetShadow();
        paragraph.AddText(" Emboss").SetEmboss();
        paragraph.AddText(" Imprint").SetImprint();
        paragraph.AddText(" plain");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"\outl Outline", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\outl0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\shad  Shadow", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\shad0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\embo  Emboss", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\embo0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\impr  Imprint", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\impr0", rtf, StringComparison.Ordinal);

        RtfParagraph readParagraph = Assert.Single(result.Document.Paragraphs);
        Assert.Contains(readParagraph.Runs, run => run.Text == "Outline" && run.Outline);
        Assert.Contains(readParagraph.Runs, run => run.Text == " Shadow" && run.Shadow);
        Assert.Contains(readParagraph.Runs, run => run.Text == " Emboss" && run.Emboss);
        Assert.Contains(readParagraph.Runs, run => run.Text == " Imprint" && run.Imprint);
        RtfRun plain = readParagraph.Runs.Single(run => run.Text == " plain");
        Assert.False(plain.Outline);
        Assert.False(plain.Shadow);
        Assert.False(plain.Emboss);
        Assert.False(plain.Imprint);
    }

    [Fact]
    public void Read_Binds_Bookmark_Markers_In_Inline_Order_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\pard {\*\bkmkstart Anchor}Bookmarked{\*\bkmkend Anchor} text\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Bookmarked text", paragraph.ToPlainText());
        Assert.Collection(paragraph.Inlines,
            inline => {
                RtfBookmarkMarker marker = Assert.IsType<RtfBookmarkMarker>(inline);
                Assert.Equal(RtfBookmarkMarkerKind.Start, marker.Kind);
                Assert.Equal("Anchor", marker.Name);
            },
            inline => Assert.Equal("Bookmarked", Assert.IsType<RtfRun>(inline).Text),
            inline => {
                RtfBookmarkMarker marker = Assert.IsType<RtfBookmarkMarker>(inline);
                Assert.Equal(RtfBookmarkMarkerKind.End, marker.Kind);
                Assert.Equal("Anchor", marker.Name);
            },
            inline => Assert.Equal(" text", Assert.IsType<RtfRun>(inline).Text));
    }

    [Fact]
    public void Write_Emits_Bookmark_Markers_And_Reads_Them_Back() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph();
        paragraph.AddBookmarkStart("Anchor");
        paragraph.AddText("Bookmarked");
        paragraph.AddBookmarkEnd("Anchor");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"{\*\bkmkstart Anchor}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\*\bkmkend Anchor}", rtf, StringComparison.Ordinal);
        RtfParagraph readParagraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Bookmarked", readParagraph.ToPlainText());
        Assert.Collection(readParagraph.Inlines,
            inline => Assert.Equal(RtfBookmarkMarkerKind.Start, Assert.IsType<RtfBookmarkMarker>(inline).Kind),
            inline => Assert.Equal("Bookmarked", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfBookmarkMarkerKind.End, Assert.IsType<RtfBookmarkMarker>(inline).Kind));
    }

    [Fact]
    public void Read_Binds_Generated_Text_Controls_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\pard Page \chpgn Section \sectnum Date \chdate Long \chdpl Short \chdpa Time \chtime\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Collection(paragraph.Inlines,
            inline => Assert.Equal("Page ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.PageNumber, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal("Section ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.SectionNumber, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal("Date ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.CurrentDate, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal("Long ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.CurrentDateLong, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal("Short ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.CurrentDateAbbreviated, Assert.IsType<RtfGeneratedText>(inline).Kind),
            inline => Assert.Equal("Time ", Assert.IsType<RtfRun>(inline).Text),
            inline => Assert.Equal(RtfGeneratedTextKind.CurrentTime, Assert.IsType<RtfGeneratedText>(inline).Kind));
    }

    [Fact]
    public void Read_Binds_Generic_Field_In_Inline_Order_With_Rich_Result_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\pard Page {\field{\*\fldinst PAGE \\* MERGEFORMAT}{\fldrslt {\b 1}}} done\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("Page 1 done", paragraph.ToPlainText());
        Assert.Collection(paragraph.Inlines,
            inline => Assert.Equal("Page ", Assert.IsType<RtfRun>(inline).Text),
            inline => {
                RtfField field = Assert.IsType<RtfField>(inline);
                Assert.Equal(@"PAGE \* MERGEFORMAT", field.Instruction);
                Assert.Equal("1", field.ToPlainText());
                Assert.Contains(field.Result.Runs, run => run.Text == "1" && run.Bold);
            },
            inline => Assert.Equal(" done", Assert.IsType<RtfRun>(inline).Text));
    }

    [Fact]
    public void Read_Binds_Hyperlink_Field_As_Field_With_Parsed_Target_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\pard See {\field{\*\fldinst HYPERLINK ""https://example.test/path"" \\o ""tip""}{\fldrslt {\b Link}}} now\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(result.Document.Paragraphs);
        Assert.Equal("See Link now", paragraph.ToPlainText());
        Assert.Collection(paragraph.Inlines,
            inline => Assert.Equal("See ", Assert.IsType<RtfRun>(inline).Text),
            inline => {
                RtfField field = Assert.IsType<RtfField>(inline);
                Assert.Equal(@"HYPERLINK ""https://example.test/path"" \o ""tip""", field.Instruction);
                Assert.Equal(new Uri("https://example.test/path"), field.Hyperlink);
                Assert.Equal("Link", field.ToPlainText());
                Assert.Contains(field.Result.Runs, run => run.Text == "Link" && run.Bold);
            },
            inline => Assert.Equal(" now", Assert.IsType<RtfRun>(inline).Text));
    }

    [Fact]
    public void Write_Emits_Generic_Field_And_Reads_It_Back() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Page ");
        RtfField field = paragraph.AddField(@"PAGE \* MERGEFORMAT");
        field.AddText("1").SetBold();

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"{\field{\*\fldinst PAGE \\* MERGEFORMAT}{\fldrslt \b 1\b0 }}", rtf, StringComparison.Ordinal);
        RtfField readField = Assert.IsType<RtfField>(Assert.Single(result.Document.Paragraphs).Inlines[1]);
        Assert.Equal(@"PAGE \* MERGEFORMAT", readField.Instruction);
        Assert.Equal("1", readField.ToPlainText());
        Assert.Contains(readField.Result.Runs, run => run.Text == "1" && run.Bold);
    }

    [Fact]
    public void Write_Emits_Hyperlink_Field_Without_Nesting_Field_Result() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("See ");
        RtfField field = paragraph.AddField(@"HYPERLINK ""https://example.test/path"" \o ""tip""");
        field.AddText("Link").SetBold();
        paragraph.AddText(" now");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"{\field{\*\fldinst HYPERLINK ""https://example.test/path"" \\o ""tip""}{\fldrslt \b Link\b0 }}", rtf, StringComparison.Ordinal);
        Assert.Equal(rtf.IndexOf(@"{\field", StringComparison.Ordinal), rtf.LastIndexOf(@"{\field", StringComparison.Ordinal));
        Assert.DoesNotContain(@"{\fldrslt {\field", rtf, StringComparison.Ordinal);
        RtfField readField = Assert.IsType<RtfField>(Assert.Single(result.Document.Paragraphs).Inlines[1]);
        Assert.Equal(new Uri("https://example.test/path"), readField.Hyperlink);
        Assert.Equal("Link", readField.ToPlainText());
        Assert.Contains(readField.Result.Runs, run => run.Text == "Link" && run.Bold);
    }

    [Fact]
    public void Read_Binds_Form_Field_Data_And_Preserves_Source_Losslessly() {
        const string rtf = @"{\rtf1\ansi\pard Name: {\field{\*\fldinst FORMTEXT}{\*\ffdata\fftype0\ffenabled1\ffownhelp1\ffownstat1\ffprot0\ffrecalc1\ffmaxlen50{\ffname Customer}{\ffdeftext Default}{\ffformat Uppercase}{\ffhelptext Help}{\ffstattext Status}{\ffentrymcr Enter}{\ffexitmcr Exit}}{\fldrslt Value}}\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(rtf, result.ToRtfLossless());
        RtfField field = Assert.IsType<RtfField>(Assert.Single(result.Document.Paragraphs).Inlines[1]);
        Assert.Equal("FORMTEXT", field.Instruction);
        Assert.Equal("Value", field.ToPlainText());
        Assert.NotNull(field.FormFieldData);
        RtfFormFieldData data = field.FormFieldData!;
        Assert.Equal(RtfFormFieldKind.Text, data.Kind);
        Assert.Equal(0, data.TypeCode);
        Assert.Equal("Customer", data.Name);
        Assert.Equal("Default", data.DefaultText);
        Assert.Equal("Uppercase", data.Format);
        Assert.Equal("Help", data.HelpText);
        Assert.Equal("Status", data.StatusText);
        Assert.Equal("Enter", data.EntryMacro);
        Assert.Equal("Exit", data.ExitMacro);
        Assert.True(data.Enabled);
        Assert.True(data.OwnHelp);
        Assert.True(data.OwnStatus);
        Assert.False(data.Protected);
        Assert.True(data.RecalculateOnExit);
        Assert.Equal(50, data.MaxLength);
        Assert.Contains(data.Controls, control => control.Name == "ffmaxlen" && control.Parameter == 50);
    }

    [Fact]
    public void Write_And_Read_DropDown_Form_Field_Data() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Choice: ");
        RtfField field = paragraph.AddField("FORMDROPDOWN");
        field.AddText("Second");
        field.SetFormFieldData(data => {
            data.Kind = RtfFormFieldKind.DropDown;
            data.Name = "Choice";
            data.Enabled = true;
            data.DefaultResult = 0;
            data.Result = 1;
            data.AddDropDownItem("First");
            data.AddDropDownItem("Second");
        });

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"{\*\ffdata\fftype2\ffenabled1\ffdefres0\ffres1{\ffname Choice}{\ffl First}{\ffl Second}}", rtf, StringComparison.Ordinal);
        RtfField readField = Assert.IsType<RtfField>(Assert.Single(result.Document.Paragraphs).Inlines[1]);
        Assert.NotNull(readField.FormFieldData);
        RtfFormFieldData data = readField.FormFieldData!;
        Assert.Equal(RtfFormFieldKind.DropDown, data.Kind);
        Assert.Equal("Choice", data.Name);
        Assert.True(data.Enabled);
        Assert.Equal(0, data.DefaultResult);
        Assert.Equal(1, data.Result);
        Assert.Equal(new[] { "First", "Second" }, data.DropDownItems);
        Assert.Equal("Second", readField.ToPlainText());
    }

    [Fact]
    public void Write_Emits_Unicode_Surrogate_Pairs_And_Reads_Them_Back() {
        RtfDocument document = RtfDocument.Create();
        document.Info.Title = "Smile 😀";
        document.AddParagraph("Body 😀");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(@"\u-10179?\u-8704?", rtf, StringComparison.Ordinal);
        Assert.Equal("Smile 😀", result.Document.Info.Title);
        Assert.Equal("Body 😀", Assert.Single(result.Document.Paragraphs).ToPlainText());
    }

    [Theory]
    [InlineData(0, @"\uc0", @"\u380  Text", "ż Text")]
    [InlineData(2, @"\uc2", @"\u380?? Text", "ż Text")]
    public void Write_Uses_Document_Unicode_Skip_Count_For_Metadata_And_Body(int skipCount, string expectedControl, string expectedEncodedText, string expectedPlainText) {
        RtfDocument document = RtfDocument.Create();
        document.Settings.SetUnicodeSkipCount(skipCount);
        document.Info.Title = "ż Text";
        document.AddParagraph("ż Text");

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Contains(expectedControl, rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\title " + expectedEncodedText + "}", rtf, StringComparison.Ordinal);
        Assert.Contains(expectedEncodedText + @"\par", rtf, StringComparison.Ordinal);
        Assert.Equal(skipCount, result.Document.Settings.UnicodeSkipCount);
        Assert.Equal(expectedPlainText, result.Document.Info.Title);
        Assert.Equal(expectedPlainText, Assert.Single(result.Document.Paragraphs).ToPlainText());
    }

    [Fact]
    public void SetUnicodeSkipCount_Rejects_Negative_Counts() {
        RtfDocument document = RtfDocument.Create();

        Assert.Throws<ArgumentOutOfRangeException>(() => document.Settings.SetUnicodeSkipCount(-1));
    }

    [Fact]
    public void Read_Honors_Uc0_When_Collecting_Metadata_Text() {
        const string rtf = @"{\rtf1\ansi{\info{\title \uc0\u380? Title}}\pard Body\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal("ż? Title", result.Document.Info.Title);
    }

    [Fact]
    public void Read_Defaults_Hex_Escapes_To_Windows1252_When_CodePage_Is_Missing() {
        const string rtf = @"{\rtf1\ansi\pard Price \'80\par}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal("Price €", Assert.Single(result.Document.Paragraphs).ToPlainText());
    }

    [Theory]
    [InlineData(874, @"\'a1\'a2", "กข")]
    [InlineData(1251, @"\'cf\'f0\'e8\'e2\'e5\'f2", "Привет")]
    [InlineData(1253, @"\'c1\'e8\'de\'ed\'e1", "Αθήνα")]
    [InlineData(1254, @"\'d0\'dd\'de\'f0\'fd\'fe", "ĞİŞğış")]
    [InlineData(1255, @"\'f9\'ec\'e5\'ed", "שלום")]
    [InlineData(1256, @"\'e3\'d1\'cd\'c8\'c7", "مرحبا")]
    [InlineData(1257, @"\'d0\'e0\'fa", "Šąś")]
    [InlineData(1258, @"\'d0\'f0\'f5\'fd", "Đđơư")]
    public void Read_Decodes_Supported_SingleByte_Ansi_CodePages_Without_Warnings(int codePage, string encodedText, string expectedText) {
        string rtf = $@"{{\rtf1\ansi\ansicpg{codePage}\pard {encodedText}\par}}";

        RtfReadResult result = RtfDocument.Read(rtf);

        Assert.Equal(expectedText, Assert.Single(result.Document.Paragraphs).ToPlainText());
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "RTF103");
    }

    [Fact]
    public void Read_Warns_When_Ansi_CodePage_Is_Unsupported_And_Uses_Documented_Fallback() {
        const string rtf = @"{\rtf1\ansi\ansicpg65001\pard Price \'80\par}";

        RtfReadResult result = RtfDocument.Read(rtf);
        RtfReadResult quiet = RtfDocument.Read(rtf, new RtfReadOptions { WarnOnUnsupportedCodePages = false });

        Assert.Equal("Price €", Assert.Single(result.Document.Paragraphs).ToPlainText());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "RTF103" && diagnostic.Severity == RtfDiagnosticSeverity.Warning);
        Assert.DoesNotContain(quiet.Diagnostics, diagnostic => diagnostic.Code == "RTF103");
    }
}
