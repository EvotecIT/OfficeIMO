using OfficeIMO.Rtf;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfHtmlDocumentSettingsTests {
    [Fact]
    public void RtfDocument_ToHtml_RoundTrips_Document_Settings_Metadata() {
        RtfDocument document = RtfDocument.Create();
        document.Settings
            .SetCharacterSet(RtfDocumentCharacterSet.Ansi, 1250)
            .SetUnicodeSkipCount(2)
            .SetDefaultFont(0)
            .SetDefaultTabWidth(720)
            .SetDefaultLanguage(1045)
            .SetDefaultFarEastLanguage(1041)
            .SetDefaultAlternateLanguage(1033)
            .SetView(kind: 4, scale: 125, zoomKind: 2, backspaceBehavior: 1)
            .SetHyphenation(automatic: false, caps: true, consecutiveLimit: 3, zoneTwips: 360)
            .SetProtection(forms: true, revisions: false, annotations: true, readOnly: false)
            .SetRevisionTracking(enabled: true, displayStyle: 3, barPlacement: 2)
            .SetDrawingGrid(horizontalSpacingTwips: 120, verticalSpacingTwips: 180, horizontalOriginTwips: 720, verticalOriginTwips: 900, horizontalShow: 2, verticalShow: 3, snapToGrid: true, useMargins: false);
        document.Settings.WidowOrphanControl = true;
        document.Settings.FacingPages = true;
        document.Settings.MirrorMargins = true;
        document.Settings.Direction = RtfTextDirection.RightToLeft;
        document.AddParagraph("Settings body");

        string html = document.ToHtml(new RtfToHtmlOptions {
            IncludeRoundTripMetadata = true,
            EmbedImagesAsDataUri = true, FragmentOnly = false, NewLine = "\n" });

        Assert.Contains("<meta name=\"officeimo-rtf-document-settings\" content=\"", html, StringComparison.Ordinal);

        RtfDocument roundTrip = html.ToRtfDocument();
        Assert.Equal(RtfDocumentCharacterSet.Ansi, roundTrip.Settings.CharacterSet);
        Assert.Equal(1250, roundTrip.Settings.AnsiCodePage);
        Assert.Equal(2, roundTrip.Settings.UnicodeSkipCount);
        Assert.Equal(0, roundTrip.Settings.DefaultFontId);
        Assert.Equal(720, roundTrip.Settings.DefaultTabWidthTwips);
        Assert.Equal(1045, roundTrip.Settings.DefaultLanguageId);
        Assert.Equal(1041, roundTrip.Settings.DefaultFarEastLanguageId);
        Assert.Equal(1033, roundTrip.Settings.DefaultAlternateLanguageId);
        Assert.Equal(4, roundTrip.Settings.ViewKind);
        Assert.Equal(125, roundTrip.Settings.ViewScale);
        Assert.Equal(2, roundTrip.Settings.ZoomKind);
        Assert.Equal(1, roundTrip.Settings.ViewBackspaceBehavior);
        Assert.True(roundTrip.Settings.WidowOrphanControl);
        Assert.False(roundTrip.Settings.AutoHyphenation);
        Assert.True(roundTrip.Settings.HyphenateCaps);
        Assert.Equal(3, roundTrip.Settings.ConsecutiveHyphenLimit);
        Assert.Equal(360, roundTrip.Settings.HyphenationZoneTwips);
        Assert.True(roundTrip.Settings.FacingPages);
        Assert.True(roundTrip.Settings.MirrorMargins);
        Assert.True(roundTrip.Settings.FormProtection);
        Assert.False(roundTrip.Settings.RevisionProtection);
        Assert.True(roundTrip.Settings.AnnotationProtection);
        Assert.False(roundTrip.Settings.ReadOnlyProtection);
        Assert.True(roundTrip.Settings.TrackRevisions);
        Assert.Equal(3, roundTrip.Settings.RevisionDisplayStyle);
        Assert.Equal(2, roundTrip.Settings.RevisionBarPlacement);
        Assert.Equal(120, roundTrip.Settings.DrawingGridHorizontalSpacingTwips);
        Assert.Equal(180, roundTrip.Settings.DrawingGridVerticalSpacingTwips);
        Assert.Equal(720, roundTrip.Settings.DrawingGridHorizontalOriginTwips);
        Assert.Equal(900, roundTrip.Settings.DrawingGridVerticalOriginTwips);
        Assert.Equal(2, roundTrip.Settings.DrawingGridHorizontalShow);
        Assert.Equal(3, roundTrip.Settings.DrawingGridVerticalShow);
        Assert.True(roundTrip.Settings.SnapToDrawingGrid);
        Assert.False(roundTrip.Settings.DrawingGridUsesMargins);
        Assert.Equal(RtfTextDirection.RightToLeft, roundTrip.Settings.Direction);
        Assert.Equal("Settings body", Assert.Single(roundTrip.Paragraphs).ToPlainText());

        string rtf = roundTrip.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        Assert.Contains(@"\ansi\ansicpg1250\deff0\uc2\deftab720\deflang1045", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\deflangfe1041\adeflang1033\viewkind4\viewscale125\viewzk2\viewbksp1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\widowctrl\hyphauto0\hyphcaps\hyphconsec3\hyphhotz360", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\facingp\margmirror\formprot\revprot0\annotprot\readprot0", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\revisions\revprop3\revbar2\dghspace120\dgvspace180\dghorigin720\dgvorigin900\dghshow2\dgvshow3\dgsnap\dgmargin0\rtldoc", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_ToRtfDocument_Document_Settings_Metadata_Does_Not_Clear_Html_Language_Direction() {
        RtfDocument metadataDocument = RtfDocument.Create();
        metadataDocument.Settings.MirrorMargins = true;
        string settingsMeta = ExtractDocumentSettingsMeta(metadataDocument.ToHtml(new RtfToHtmlOptions {
            IncludeRoundTripMetadata = true,
            EmbedImagesAsDataUri = true, FragmentOnly = false }));
        string html = "<html lang=\"ar-SA\" dir=\"rtl\"><head>" + settingsMeta + "</head><body><p>Body</p></body></html>";

        RtfDocument document = html.ToRtfDocument();

        Assert.Equal(1025, document.Settings.DefaultLanguageId);
        Assert.Equal(RtfTextDirection.RightToLeft, document.Settings.Direction);
        Assert.True(document.Settings.MirrorMargins);
    }

    private static string ExtractDocumentSettingsMeta(string html) {
        const string prefix = "<meta name=\"officeimo-rtf-document-settings\"";
        int start = html.IndexOf(prefix, StringComparison.Ordinal);
        Assert.True(start >= 0);
        int end = html.IndexOf("\">", start, StringComparison.Ordinal);
        Assert.True(end >= start);
        return html.Substring(start, end - start + 2);
    }
}
