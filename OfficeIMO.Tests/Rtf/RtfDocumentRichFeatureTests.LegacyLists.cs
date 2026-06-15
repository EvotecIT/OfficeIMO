using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class RtfDocumentRichFeatureTests {
    [Fact]
    public void Read_Binds_Word95_Legacy_Numbering_Destination_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\pntext\f0 3.\tab}{\*\pn\pnlvlbody\pndec\pnstart3\pnindent720\pnsp240\pnprev\pnql{\pntxtb (}{\pntxta )}}\pard\li720\fi-360 Legacy\par}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(read.Document.Paragraphs);
        Assert.Equal("Legacy", paragraph.ToPlainText());
        Assert.Equal(RtfListKind.Decimal, paragraph.ListKind);
        Assert.True(paragraph.LegacyNumbering.Enabled);
        Assert.Equal(RtfLegacyNumberingLevelKind.Body, paragraph.LegacyNumbering.LevelKind);
        Assert.Equal(RtfLegacyNumberingStyle.Decimal, paragraph.LegacyNumbering.NumberStyle);
        Assert.Equal(3, paragraph.LegacyNumbering.StartAt);
        Assert.Equal(720, paragraph.LegacyNumbering.IndentTwips);
        Assert.Equal(240, paragraph.LegacyNumbering.SpaceTwips);
        Assert.True(paragraph.LegacyNumbering.IncludePreviousLevels);
        Assert.Equal(RtfLegacyNumberingAlignment.Left, paragraph.LegacyNumbering.Alignment);
        Assert.Equal("(", paragraph.LegacyNumbering.TextBefore);
        Assert.Equal(")", paragraph.LegacyNumbering.TextAfter);
        Assert.Equal(720, paragraph.LeftIndentTwips);
        Assert.Equal(-360, paragraph.FirstLineIndentTwips);
    }

    [Fact]
    public void Write_And_Read_Word95_Legacy_Numbering_Destination() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Legacy");
        paragraph.SetLegacyNumbering(numbering => {
            numbering.LevelKind = RtfLegacyNumberingLevelKind.Body;
            numbering.NumberStyle = RtfLegacyNumberingStyle.Decimal;
            numbering.FontId = 1;
            numbering.FontSizeHalfPoints = 24;
            numbering.Bold = true;
            numbering.Italic = false;
            numbering.AllCaps = true;
            numbering.SmallCaps = false;
            numbering.UnderlineStyle = RtfUnderlineStyle.Double;
            numbering.Strike = true;
            numbering.ForegroundColorIndex = 2;
            numbering.TextBefore = "(";
            numbering.TextAfter = ")";
            numbering.NumberEachCellOnce = true;
            numbering.NumberAcrossRows = false;
            numbering.IndentTwips = 720;
            numbering.SpaceTwips = 240;
            numbering.IncludePreviousLevels = true;
            numbering.Alignment = RtfLegacyNumberingAlignment.Center;
            numbering.StartAt = 3;
            numbering.HangingIndent = true;
            numbering.RestartAfterSection = false;
        });
        paragraph.SetIndentation(leftTwips: 720, firstLineTwips: -360);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\*\pn\pnlvlbody\pndec\pnf1\pnfs24\pnb1\pni0\pncaps1\pnscaps0\pnuldb\pnstrike1\pncf2{\pntxtb (}{\pntxta )}\pnnumonce1\pnacross0\pnindent720\pnsp240\pnprev1\pnqc\pnstart3\pnhang1\pnrestart0}", rtf, StringComparison.Ordinal);
        RtfParagraph roundTrip = Assert.Single(read.Document.Paragraphs);
        Assert.True(roundTrip.LegacyNumbering.Enabled);
        Assert.Equal(RtfLegacyNumberingLevelKind.Body, roundTrip.LegacyNumbering.LevelKind);
        Assert.Equal(RtfLegacyNumberingStyle.Decimal, roundTrip.LegacyNumbering.NumberStyle);
        Assert.Equal(1, roundTrip.LegacyNumbering.FontId);
        Assert.Equal(24, roundTrip.LegacyNumbering.FontSizeHalfPoints);
        Assert.True(roundTrip.LegacyNumbering.Bold);
        Assert.False(roundTrip.LegacyNumbering.Italic);
        Assert.True(roundTrip.LegacyNumbering.AllCaps);
        Assert.False(roundTrip.LegacyNumbering.SmallCaps);
        Assert.Equal(RtfUnderlineStyle.Double, roundTrip.LegacyNumbering.UnderlineStyle);
        Assert.True(roundTrip.LegacyNumbering.Strike);
        Assert.Equal(2, roundTrip.LegacyNumbering.ForegroundColorIndex);
        Assert.Equal("(", roundTrip.LegacyNumbering.TextBefore);
        Assert.Equal(")", roundTrip.LegacyNumbering.TextAfter);
        Assert.True(roundTrip.LegacyNumbering.NumberEachCellOnce);
        Assert.False(roundTrip.LegacyNumbering.NumberAcrossRows);
        Assert.Equal(720, roundTrip.LegacyNumbering.IndentTwips);
        Assert.Equal(240, roundTrip.LegacyNumbering.SpaceTwips);
        Assert.True(roundTrip.LegacyNumbering.IncludePreviousLevels);
        Assert.Equal(RtfLegacyNumberingAlignment.Center, roundTrip.LegacyNumbering.Alignment);
        Assert.Equal(3, roundTrip.LegacyNumbering.StartAt);
        Assert.True(roundTrip.LegacyNumbering.HangingIndent);
        Assert.False(roundTrip.LegacyNumbering.RestartAfterSection);
    }

    [Fact]
    public void Read_Binds_Word95_Legacy_Numbering_In_Stylesheet_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\stylesheet{\s1{\*\pn\pnlvlbody\pnucrm\pnstart4\pnindent720\pnsp240\pnqr{\pntxtb Chapter }{\pntxta .}}\li720\fi-360 Numbered Heading;}}\pard\s1 Heading\par}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfStyle style = Assert.Single(read.Document.Styles);
        Assert.Equal(1, style.Id);
        Assert.Equal("Numbered Heading", style.Name);
        Assert.Equal(720, style.LeftIndentTwips);
        Assert.Equal(-360, style.FirstLineIndentTwips);
        Assert.True(style.LegacyNumbering.Enabled);
        Assert.Equal(RtfLegacyNumberingLevelKind.Body, style.LegacyNumbering.LevelKind);
        Assert.Equal(RtfLegacyNumberingStyle.UpperRoman, style.LegacyNumbering.NumberStyle);
        Assert.Equal(4, style.LegacyNumbering.StartAt);
        Assert.Equal(720, style.LegacyNumbering.IndentTwips);
        Assert.Equal(240, style.LegacyNumbering.SpaceTwips);
        Assert.Equal(RtfLegacyNumberingAlignment.Right, style.LegacyNumbering.Alignment);
        Assert.Equal("Chapter ", style.LegacyNumbering.TextBefore);
        Assert.Equal(".", style.LegacyNumbering.TextAfter);
    }

    [Fact]
    public void Write_And_Read_Word95_Legacy_Numbering_In_Stylesheet() {
        RtfDocument document = RtfDocument.Create();
        RtfStyle style = document.AddStyle(1, "Numbered Heading");
        style.LeftIndentTwips = 720;
        style.FirstLineIndentTwips = -360;
        style.SetLegacyNumbering(numbering => {
            numbering.LevelKind = RtfLegacyNumberingLevelKind.Body;
            numbering.NumberStyle = RtfLegacyNumberingStyle.UpperRoman;
            numbering.StartAt = 4;
            numbering.IndentTwips = 720;
            numbering.SpaceTwips = 240;
            numbering.Alignment = RtfLegacyNumberingAlignment.Right;
            numbering.TextBefore = "Chapter";
            numbering.TextAfter = ".";
        });
        document.AddParagraph("Heading").SetStyle(1);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\stylesheet{\s1{\*\pn\pnlvlbody\pnucrm{\pntxtb Chapter}{\pntxta .}\pnindent720\pnsp240\pnqr\pnstart4}\li720\fi-360 Numbered Heading;}}", rtf, StringComparison.Ordinal);
        RtfStyle roundTrip = Assert.Single(read.Document.Styles);
        Assert.True(roundTrip.LegacyNumbering.Enabled);
        Assert.Equal(RtfLegacyNumberingLevelKind.Body, roundTrip.LegacyNumbering.LevelKind);
        Assert.Equal(RtfLegacyNumberingStyle.UpperRoman, roundTrip.LegacyNumbering.NumberStyle);
        Assert.Equal(4, roundTrip.LegacyNumbering.StartAt);
        Assert.Equal(720, roundTrip.LegacyNumbering.IndentTwips);
        Assert.Equal(240, roundTrip.LegacyNumbering.SpaceTwips);
        Assert.Equal(RtfLegacyNumberingAlignment.Right, roundTrip.LegacyNumbering.Alignment);
        Assert.Equal("Chapter", roundTrip.LegacyNumbering.TextBefore);
        Assert.Equal(".", roundTrip.LegacyNumbering.TextAfter);
    }

    [Fact]
    public void Read_Binds_Word97_ListText_As_Marker_Not_Body_Text_Losslessly() {
        const string rtf = @"{\rtf1\ansi{\listtext\f0 1.\tab}\pard\ls3\ilvl0 Item\par}";

        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Equal(rtf, read.ToRtfLossless());
        RtfParagraph paragraph = Assert.Single(read.Document.Paragraphs);
        Assert.Equal("Item", paragraph.ToPlainText());
        Assert.NotNull(paragraph.ListText);
        Assert.Equal("1.\t", paragraph.ListText!.ToPlainText());
        Assert.Equal(3, paragraph.ListId);
        Assert.Equal(0, paragraph.ListLevel);
    }

    [Fact]
    public void Write_And_Read_Word97_ListText_Marker() {
        RtfDocument document = RtfDocument.Create();
        RtfParagraph paragraph = document.AddParagraph("Item");
        paragraph.SetList(listId: 3, level: 0, kind: RtfListKind.Decimal);
        paragraph.SetListText(marker => marker.AddText("1.\t").FontId = 0);

        string rtf = document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        RtfReadResult read = RtfDocument.Read(rtf);

        Assert.Contains(@"{\listtext \f0 1.\tab }", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pard\pn\pnlvlbody\ls3\ilvl0\ql Item\par", rtf, StringComparison.Ordinal);
        RtfParagraph roundTrip = Assert.Single(read.Document.Paragraphs);
        Assert.Equal("Item", roundTrip.ToPlainText());
        Assert.NotNull(roundTrip.ListText);
        Assert.Equal("1.\t", roundTrip.ListText!.ToPlainText());
        Assert.Equal(3, roundTrip.ListId);
        Assert.Equal(0, roundTrip.ListLevel);
    }
}
