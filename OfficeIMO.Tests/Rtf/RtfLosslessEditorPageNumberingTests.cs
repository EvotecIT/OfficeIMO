using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfLosslessEditorPageNumberingTests {
    [Fact]
    public void SetPageNumbering_Replaces_Duplicates_And_Preserves_Metadata_And_Body() {
        const string rtf = @"{\rtf1\ansi\pgnstarts1\pgnstarts2\pgnrestart\pgncont\pgnx100\pgny200\pgndec\pgnucrm{\info{\title Keep}}\pard Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetPageNumbering(
            start: 5,
            restart: true,
            format: RtfPageNumberFormat.LowerRoman,
            positionXTwips: 720,
            positionYTwips: 900);

        const string expected = @"{\rtf1\ansi\pgnstarts5\pgnrestart\pgnx720\pgny900\pgnlcrm{\info{\title Keep}}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        RtfPageSetup pageSetup = read.Document.PageSetup;
        Assert.Equal(5, pageSetup.PageNumberStart);
        Assert.True(pageSetup.PageNumberRestart);
        Assert.Equal(720, pageSetup.PageNumberPositionXTwips);
        Assert.Equal(900, pageSetup.PageNumberPositionYTwips);
        Assert.Equal(RtfPageNumberFormat.LowerRoman, pageSetup.PageNumberFormat);
        Assert.Equal("Keep", read.Document.Info.Title);
        Assert.Contains(@"Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetPageNumbering_Creates_Continuous_Numbering_And_Removes_All_Page_Numbering_Controls() {
        const string rtf = @"{\rtf1\ansi{\info{\title Keep}}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetPageNumbering(
            start: 3,
            restart: false,
            format: RtfPageNumberFormat.UpperLetter,
            positionXTwips: 360,
            positionYTwips: 480);

        const string expected = @"{\rtf1\ansi\pgnstarts3\pgncont\pgnx360\pgny480\pgnucltr{\info{\title Keep}}\pard Body\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfPageSetup pageSetup = editor.ToReadResult().Document.PageSetup;
        Assert.Equal(3, pageSetup.PageNumberStart);
        Assert.False(pageSetup.PageNumberRestart);
        Assert.Equal(360, pageSetup.PageNumberPositionXTwips);
        Assert.Equal(480, pageSetup.PageNumberPositionYTwips);
        Assert.Equal(RtfPageNumberFormat.UpperLetter, pageSetup.PageNumberFormat);

        editor.SetPageNumbering();

        Assert.Equal(rtf, editor.ToRtf());
        pageSetup = editor.ToReadResult().Document.PageSetup;
        Assert.Null(pageSetup.PageNumberStart);
        Assert.Null(pageSetup.PageNumberRestart);
        Assert.Null(pageSetup.PageNumberPositionXTwips);
        Assert.Null(pageSetup.PageNumberPositionYTwips);
        Assert.Null(pageSetup.PageNumberFormat);
    }
}
