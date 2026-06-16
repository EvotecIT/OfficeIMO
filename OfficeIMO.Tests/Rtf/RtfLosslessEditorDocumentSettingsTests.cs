using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfLosslessEditorDocumentSettingsTests {
    [Fact]
    public void SetDocumentSettings_Replaces_Duplicates_And_Preserves_Metadata_And_Body() {
        const string rtf = @"{\rtf1\ansi\uc2\deftab360\deftab720\deflang1045\deflangfe1041\adeflang1033\viewkind4\viewscale125\viewzk2\viewbksp1\widowctrl\hyphauto0\hyphcaps\hyphconsec3\hyphhotz360\facingp\margmirror\formprot\revprot0\annotprot\readprot0\revisions\revprop3\revbar2\dghspace120\dgvspace180\dghorigin720\dgvorigin900\dghshow2\dgvshow3\dgsnap\dgmargin0\rtldoc{\info{\title Keep}}\pard Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetDefaultTabWidth(1440);
        editor.SetDefaultLanguages(defaultLanguageId: 1033, defaultAlternateLanguageId: 1045);
        editor.SetDocumentView(kind: 5, zoomKind: 3);
        editor.SetDocumentLayoutOptions(widowOrphanControl: false, mirrorMargins: true);
        editor.SetDocumentHyphenation(automatic: true, caps: false, zoneTwips: 480);
        editor.SetDocumentProtection(forms: false, revisions: true, readOnly: true);
        editor.SetRevisionTracking(enabled: false, displayStyle: 4);
        editor.SetDrawingGrid(horizontalSpacingTwips: 240, verticalOriginTwips: 960, verticalShow: 4, snapToGrid: false, useMargins: true);
        editor.SetDocumentDirection(RtfTextDirection.LeftToRight);

        const string expected = @"{\rtf1\ansi\uc2\deftab1440\deflang1033\adeflang1045\viewkind5\viewzk3\widowctrl0\hyphauto\hyphcaps0\hyphhotz480\margmirror\formprot0\revprot\readprot\revisions0\revprop4\dghspace240\dgvorigin960\dgvshow4\dgsnap0\dgmargin\ltrdoc{\info{\title Keep}}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        RtfDocumentSettings settings = read.Document.Settings;
        Assert.Equal(2, settings.UnicodeSkipCount);
        Assert.Equal(1440, settings.DefaultTabWidthTwips);
        Assert.Equal(1033, settings.DefaultLanguageId);
        Assert.Null(settings.DefaultFarEastLanguageId);
        Assert.Equal(1045, settings.DefaultAlternateLanguageId);
        Assert.Equal(5, settings.ViewKind);
        Assert.Null(settings.ViewScale);
        Assert.Equal(3, settings.ZoomKind);
        Assert.Null(settings.ViewBackspaceBehavior);
        Assert.False(settings.WidowOrphanControl);
        Assert.True(settings.AutoHyphenation);
        Assert.False(settings.HyphenateCaps);
        Assert.Null(settings.ConsecutiveHyphenLimit);
        Assert.Equal(480, settings.HyphenationZoneTwips);
        Assert.Null(settings.FacingPages);
        Assert.True(settings.MirrorMargins);
        Assert.False(settings.FormProtection);
        Assert.True(settings.RevisionProtection);
        Assert.Null(settings.AnnotationProtection);
        Assert.True(settings.ReadOnlyProtection);
        Assert.False(settings.TrackRevisions);
        Assert.Equal(4, settings.RevisionDisplayStyle);
        Assert.Null(settings.RevisionBarPlacement);
        Assert.Equal(240, settings.DrawingGridHorizontalSpacingTwips);
        Assert.Null(settings.DrawingGridVerticalSpacingTwips);
        Assert.Null(settings.DrawingGridHorizontalOriginTwips);
        Assert.Equal(960, settings.DrawingGridVerticalOriginTwips);
        Assert.Null(settings.DrawingGridHorizontalShow);
        Assert.Equal(4, settings.DrawingGridVerticalShow);
        Assert.False(settings.SnapToDrawingGrid);
        Assert.True(settings.DrawingGridUsesMargins);
        Assert.Equal(RtfTextDirection.LeftToRight, settings.Direction);
        Assert.Equal("Keep", read.Document.Info.Title);
        Assert.Contains(@"Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetDocumentSettings_Creates_And_Removes_Settings_Before_Metadata() {
        const string rtf = @"{\rtf1\ansi{\info{\title Keep}}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetDefaultTabWidth(720);
        editor.SetDefaultLanguages(defaultLanguageId: 1045, defaultFarEastLanguageId: 1041, defaultAlternateLanguageId: 1033);
        editor.SetDocumentView(kind: 4, scale: 125, zoomKind: 2, backspaceBehavior: 1);
        editor.SetDocumentLayoutOptions(widowOrphanControl: true, facingPages: true, mirrorMargins: false);
        editor.SetDocumentHyphenation(automatic: false, caps: true, consecutiveLimit: 3, zoneTwips: 360);
        editor.SetDocumentProtection(forms: true, revisions: false, annotations: true, readOnly: false);
        editor.SetRevisionTracking(enabled: true, displayStyle: 3, barPlacement: 2);
        editor.SetDrawingGrid(
            horizontalSpacingTwips: 120,
            verticalSpacingTwips: 180,
            horizontalOriginTwips: 720,
            verticalOriginTwips: 900,
            horizontalShow: 2,
            verticalShow: 3,
            snapToGrid: true,
            useMargins: false);
        editor.SetDocumentDirection(RtfTextDirection.RightToLeft);

        const string expected = @"{\rtf1\ansi\deftab720\deflang1045\deflangfe1041\adeflang1033\viewkind4\viewscale125\viewzk2\viewbksp1\widowctrl\hyphauto0\hyphcaps\hyphconsec3\hyphhotz360\facingp\margmirror0\formprot\revprot0\annotprot\readprot0\revisions\revprop3\revbar2\dghspace120\dgvspace180\dghorigin720\dgvorigin900\dghshow2\dgvshow3\dgsnap\dgmargin0\rtldoc{\info{\title Keep}}\pard Body\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfDocumentSettings settings = editor.ToReadResult().Document.Settings;
        Assert.Equal(720, settings.DefaultTabWidthTwips);
        Assert.Equal(1045, settings.DefaultLanguageId);
        Assert.True(settings.WidowOrphanControl);
        Assert.False(settings.AutoHyphenation);
        Assert.True(settings.HyphenateCaps);
        Assert.True(settings.FacingPages);
        Assert.False(settings.MirrorMargins);
        Assert.True(settings.FormProtection);
        Assert.False(settings.RevisionProtection);
        Assert.True(settings.AnnotationProtection);
        Assert.False(settings.ReadOnlyProtection);
        Assert.True(settings.TrackRevisions);
        Assert.False(settings.DrawingGridUsesMargins);
        Assert.Equal(RtfTextDirection.RightToLeft, settings.Direction);

        editor.SetDefaultTabWidth(null);
        editor.SetDefaultLanguages();
        editor.SetDocumentView();
        editor.SetDocumentLayoutOptions();
        editor.SetDocumentHyphenation();
        editor.SetDocumentProtection();
        editor.SetRevisionTracking();
        editor.SetDrawingGrid();
        editor.SetDocumentDirection(null);

        Assert.Equal(rtf, editor.ToRtf());
    }
}
