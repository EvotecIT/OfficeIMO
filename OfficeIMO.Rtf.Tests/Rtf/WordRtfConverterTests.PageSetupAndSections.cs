using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using System.Linq;
using DocumentFormat.OpenXml.CustomProperties;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public partial class WordRtfConverterTests {
    [Fact]
    public void Word_To_Rtf_Bridge_Carries_PageSetup_Headers_And_Footers() {
        using WordDocument word = WordDocument.Create();
        word.PageOrientation = PageOrientationValues.Landscape;
        word.PageSettings.Width = 16838U;
        word.PageSettings.Height = 11906U;
        word.Margins.Left = 1440U;
        word.Margins.Right = 720U;
        word.Margins.Top = 1080;
        word.Margins.Bottom = 1080;
        word.Margins.Gutter = 180U;
        word.Margins.HeaderDistance = 360U;
        word.Margins.FooterDistance = 540U;
        word.RtlGutter = true;
        word.Settings.MirrorMargins = true;
        word.AddPageNumbering(5, NumberFormatValues.LowerRoman);
        word.Borders.TopStyle = BorderValues.Single;
        word.Borders.TopSize = 12U;
        word.Borders.TopSpace = 24U;
        word.Borders.TopColorHex = "FF0000";
        word.Borders.BottomStyle = BorderValues.Double;
        word.Borders.BottomSize = 18U;
        word.Borders.BottomSpace = 30U;
        word.Borders.BottomColorHex = "0000FF";
        PageBorders pageBorders = word.Sections[0]._sectionProperties.GetFirstChild<PageBorders>()!;
        pageBorders.Display = PageBorderDisplayValues.NotFirstPage;
        pageBorders.ZOrder = PageBorderZOrderValues.Front;
        pageBorders.OffsetFrom = PageBorderOffsetValues.Page;
        word.DifferentFirstPage = true;
        word.DifferentOddAndEvenPages = true;
        word.HeaderDefaultOrCreate.AddParagraph().AddText("Header ").AddText("bold").SetBold();
        word.HeaderFirstOrCreate.AddParagraph("First header");
        word.HeaderEvenOrCreate.AddParagraph("Even header");
        word.FooterDefaultOrCreate.AddParagraph("Footer text");
        word.FooterEvenOrCreate.AddParagraph("Even footer");
        word.AddParagraph("Body");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });

        Assert.Equal(16838, rtfDocument.PageSetup.PaperWidthTwips);
        Assert.Equal(11906, rtfDocument.PageSetup.PaperHeightTwips);
        Assert.Equal(1440, rtfDocument.PageSetup.MarginLeftTwips);
        Assert.Equal(720, rtfDocument.PageSetup.MarginRightTwips);
        Assert.Equal(1080, rtfDocument.PageSetup.MarginTopTwips);
        Assert.Equal(1080, rtfDocument.PageSetup.MarginBottomTwips);
        Assert.Equal(180, rtfDocument.PageSetup.GutterWidthTwips);
        Assert.Equal(360, rtfDocument.PageSetup.HeaderDistanceTwips);
        Assert.Equal(540, rtfDocument.PageSetup.FooterDistanceTwips);
        Assert.True(rtfDocument.PageSetup.RtlGutter);
        Assert.Equal(5, rtfDocument.PageSetup.PageNumberStart);
        Assert.True(rtfDocument.PageSetup.PageNumberRestart);
        Assert.Equal(RtfPageNumberFormat.LowerRoman, rtfDocument.PageSetup.PageNumberFormat);
        Assert.Equal(RtfPageBorderScope.AllExceptFirstPageInSection, rtfDocument.PageSetup.PageBorders.Scope);
        Assert.False(rtfDocument.PageSetup.PageBorders.DisplayBehindText);
        Assert.Equal(RtfPageBorderOffset.PageEdge, rtfDocument.PageSetup.PageBorders.OffsetFrom);
        Assert.Equal(RtfPageBorderStyle.Single, rtfDocument.PageSetup.PageBorders.Top.Style);
        Assert.Equal(12, rtfDocument.PageSetup.PageBorders.Top.Width);
        Assert.Equal(24, rtfDocument.PageSetup.PageBorders.Top.Space);
        Assert.Equal(1, rtfDocument.PageSetup.PageBorders.Top.ColorIndex);
        Assert.Equal(RtfPageBorderStyle.Double, rtfDocument.PageSetup.PageBorders.Bottom.Style);
        Assert.Equal(18, rtfDocument.PageSetup.PageBorders.Bottom.Width);
        Assert.Equal(30, rtfDocument.PageSetup.PageBorders.Bottom.Space);
        Assert.Equal(2, rtfDocument.PageSetup.PageBorders.Bottom.ColorIndex);
        Assert.True(rtfDocument.PageSetup.Landscape);
        Assert.True(rtfDocument.PageSetup.DifferentFirstPageHeaderFooter);
        Assert.True(rtfDocument.Settings.FacingPages);
        Assert.True(rtfDocument.Settings.MirrorMargins);
        Assert.Contains(rtfDocument.HeaderFooters, item => item.Kind == RtfHeaderFooterKind.Header && item.ToPlainText() == "Header bold");
        Assert.Contains(rtfDocument.HeaderFooters, item => item.Kind == RtfHeaderFooterKind.FirstHeader && item.ToPlainText() == "First header");
        Assert.Contains(rtfDocument.HeaderFooters, item => item.Kind == RtfHeaderFooterKind.LeftHeader && item.ToPlainText() == "Even header");
        Assert.Contains(rtfDocument.HeaderFooters, item => item.Kind == RtfHeaderFooterKind.Footer && item.ToPlainText() == "Footer text");
        Assert.Contains(rtfDocument.HeaderFooters, item => item.Kind == RtfHeaderFooterKind.LeftFooter && item.ToPlainText() == "Even footer");
        Assert.Contains(@"\paperw16838", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\landscape", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\titlepg", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\facingp", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\margmirror", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\gutter180", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\headery360", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\footery540", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\rtlgutter", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pgnstarts5", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pgnrestart", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pgnlcrm", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pgbrdropt34", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pgbrdrt\brdrs\brdrw12\brsp24\brdrcf1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\pgbrdrb\brdrdb\brdrw18\brsp30\brdrcf2", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\header", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\headerf", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\headerl", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\footer", rtf, StringComparison.Ordinal);
        Assert.Contains(@"{\footerl", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Treats_Mirrored_Margin_Preset_As_Mirror_Margins() {
        using WordDocument word = WordDocument.Create();
        word.Margins.Type = WordMargin.Mirrored;
        word.AddParagraph("Body");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });

        Assert.True(rtfDocument.Settings.MirrorMargins);
        Assert.Contains(@"\margmirror", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Word_To_Rtf_Bridge_Carries_Document_Settings() {
        using WordDocument word = WordDocument.Create();
        word.Settings.DefaultTabStop = 1440;
        word.Settings.ZoomPercentage = 125;
        Settings settings = word._wordprocessingDocument.MainDocumentPart!.DocumentSettingsPart!.Settings!;
        settings.RemoveAllChildren<DocumentProtection>();
        settings.Append(new DocumentProtection { Edit = DocumentProtectionValues.ReadOnly });
        word.AddParagraph("Body");

        RtfDocument rtfDocument = word.ToRtfDocument();
        string rtf = word.ToRtf(new RtfWriteOptions { IncludeGenerator = false });

        Assert.Equal(1440, rtfDocument.Settings.DefaultTabWidthTwips);
        Assert.Equal(125, rtfDocument.Settings.ViewScale);
        Assert.True(rtfDocument.Settings.ReadOnlyProtection);
        Assert.Contains(@"\deftab1440", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\viewscale125", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\readprot", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Rtf_To_Word_Bridge_Carries_Document_Settings() {
        RtfDocument rtfDocument = RtfDocument.Create();
        rtfDocument.Settings
            .SetDefaultTabWidth(1440)
            .SetView(scale: 125)
            .SetProtection(readOnly: true);
        rtfDocument.Settings.FacingPages = true;
        rtfDocument.Settings.MirrorMargins = true;
        rtfDocument.AddParagraph("Body");

        using WordDocument word = rtfDocument.ToWordDocument();
        RtfDocument roundTrip = word.ToRtfDocument();

        Assert.Equal(1440, word.Settings.DefaultTabStop);
        Assert.Equal(125, word.Settings.ZoomPercentage);
        Assert.Equal(DocumentProtectionValues.ReadOnly, word.Settings.ProtectionType);
        Assert.True(word.DifferentOddAndEvenPages);
        Assert.True(word.Settings.MirrorMargins);
        Assert.Equal(1440, roundTrip.Settings.DefaultTabWidthTwips);
        Assert.Equal(125, roundTrip.Settings.ViewScale);
        Assert.True(roundTrip.Settings.ReadOnlyProtection);
        Assert.True(roundTrip.Settings.FacingPages);
        Assert.True(roundTrip.Settings.MirrorMargins);
    }
}
