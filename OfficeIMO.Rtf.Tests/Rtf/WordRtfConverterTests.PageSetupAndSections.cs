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
        word.PageOrientation = OfficePageOrientation.Landscape;
        word.PageSettings.Width = global::OfficeIMO.OfficeLength.Points(841.9D);
        word.PageSettings.Height = global::OfficeIMO.OfficeLength.Points(595.3D);
        word.Margins.Left = global::OfficeIMO.OfficeLength.Points(72D);
        word.Margins.Right = global::OfficeIMO.OfficeLength.Points(36D);
        word.Margins.Top = global::OfficeIMO.OfficeLength.Points(54D);
        word.Margins.Bottom = global::OfficeIMO.OfficeLength.Points(54D);
        word.Margins.Gutter = global::OfficeIMO.OfficeLength.Points(9D);
        word.Margins.HeaderDistance = global::OfficeIMO.OfficeLength.Points(18D);
        word.Margins.FooterDistance = global::OfficeIMO.OfficeLength.Points(27D);
        word.RtlGutter = true;
        word.Settings.MirrorMargins = true;
        word.AddPageNumbering(5, WordNumberFormat.LowerRoman);
        word.Borders.TopStyle = WordBorderStyle.Single;
        word.Borders.TopSize = OfficeLength.Points(1.5);
        word.Borders.TopSpace = OfficeLength.Points(24);
        word.Borders.TopColorHex = "FF0000";
        word.Borders.BottomStyle = WordBorderStyle.Double;
        word.Borders.BottomSize = OfficeLength.Points(2.25);
        word.Borders.BottomSpace = OfficeLength.Points(30);
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
        Assert.Equal(30, rtfDocument.PageSetup.PageBorders.Top.Width);
        Assert.Equal(24, rtfDocument.PageSetup.PageBorders.Top.Space);
        Assert.Equal("#FF0000", rtfDocument.GetColor(rtfDocument.PageSetup.PageBorders.Top.ColorIndex!.Value)!.ToString());
        Assert.Equal(RtfPageBorderStyle.Double, rtfDocument.PageSetup.PageBorders.Bottom.Style);
        Assert.Equal(45, rtfDocument.PageSetup.PageBorders.Bottom.Width);
        Assert.Equal(30, rtfDocument.PageSetup.PageBorders.Bottom.Space);
        Assert.Equal("#0000FF", rtfDocument.GetColor(rtfDocument.PageSetup.PageBorders.Bottom.ColorIndex!.Value)!.ToString());
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
        Assert.Contains($@"\pgbrdrt\brdrs\brdrw30\brsp24\brdrcf{rtfDocument.PageSetup.PageBorders.Top.ColorIndex}", rtf, StringComparison.Ordinal);
        Assert.Contains($@"\pgbrdrb\brdrdb\brdrw45\brsp30\brdrcf{rtfDocument.PageSetup.PageBorders.Bottom.ColorIndex}", rtf, StringComparison.Ordinal);
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
        Assert.Equal(WordDocumentProtectionType.ReadOnly, word.Settings.ProtectionType);
        // The selection flag survives without inventing an empty even-page story.
        Assert.False(word.DifferentOddAndEvenPages);
        Assert.Null(word.Sections[0].Header.Even);
        Assert.Null(word.Sections[0].Footer.Even);
        Assert.True(word.Settings.MirrorMargins);
        Assert.Equal(1440, roundTrip.Settings.DefaultTabWidthTwips);
        Assert.Equal(125, roundTrip.Settings.ViewScale);
        Assert.True(roundTrip.Settings.ReadOnlyProtection);
        Assert.True(roundTrip.Settings.FacingPages);
        Assert.True(roundTrip.Settings.MirrorMargins);
    }
}
