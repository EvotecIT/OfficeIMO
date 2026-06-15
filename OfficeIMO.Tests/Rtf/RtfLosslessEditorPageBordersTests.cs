using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfLosslessEditorPageBordersTests {
    [Fact]
    public void SetPageBorders_Replaces_Duplicates_And_Preserves_Metadata_And_Body() {
        const string rtf = @"{\rtf1\ansi\pgbrdrhead\pgbrdrfoot\pgbrdropt34\pgbrdrsnap\pgbrdrt\brdrs\brdrw12\brsp24\brdrcf1\pgbrdrt\brdrdb\brdrw18\pgbrdrb\brdrdot\brsp20{\info{\title Keep}}\pard Body \'80\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetPageBorderOptions(
            includeHeader: false,
            includeFooter: true,
            snapToPageBorder: false,
            scope: RtfPageBorderScope.WholeDocument,
            displayBehindText: true,
            offsetFrom: RtfPageBorderOffset.PageEdge);
        editor.SetPageBorder(
            RtfPageBorderSide.Top,
            RtfPageBorderStyle.Dotted,
            width: 8,
            space: 12,
            colorIndex: 2,
            frame: true);
        editor.RemovePageBorder(RtfPageBorderSide.Bottom);

        const string expected = @"{\rtf1\ansi\pgbrdrfoot\pgbrdropt43\pgbrdrt\brdrdot\brdrw8\brsp12\brdrcf2\brdrframe{\info{\title Keep}}\pard Body \'80\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfReadResult read = editor.ToReadResult();
        RtfPageBorders borders = read.Document.PageSetup.PageBorders;
        Assert.False(borders.IncludeHeader);
        Assert.True(borders.IncludeFooter);
        Assert.False(borders.SnapToPageBorder);
        Assert.Equal(RtfPageBorderScope.WholeDocument, borders.Scope);
        Assert.True(borders.DisplayBehindText);
        Assert.Equal(RtfPageBorderOffset.PageEdge, borders.OffsetFrom);
        Assert.Equal(RtfPageBorderStyle.Dotted, borders.Top.Style);
        Assert.Equal(8, borders.Top.Width);
        Assert.Equal(12, borders.Top.Space);
        Assert.Equal(2, borders.Top.ColorIndex);
        Assert.True(borders.Top.Frame);
        Assert.False(borders.Bottom.HasAnyValue);
        Assert.Equal("Keep", read.Document.Info.Title);
        Assert.Contains(@"Body \'80", editor.ToRtf(), StringComparison.Ordinal);
    }

    [Fact]
    public void SetPageBorders_Creates_And_Removes_Options_And_Sides_Before_Metadata() {
        const string rtf = @"{\rtf1\ansi{\info{\title Keep}}\pard Body\par}";

        RtfLosslessEditor editor = RtfDocument.Read(rtf).EditLossless();
        editor.SetPageBorderOptions(
            includeHeader: true,
            includeFooter: true,
            snapToPageBorder: true,
            scope: RtfPageBorderScope.FirstPageInSection,
            displayBehindText: false,
            offsetFrom: RtfPageBorderOffset.Text);
        editor.SetPageBorder(RtfPageBorderSide.Top, RtfPageBorderStyle.Single, width: 12, space: 24, colorIndex: 1);
        editor.SetPageBorder(RtfPageBorderSide.Bottom, RtfPageBorderStyle.Double, width: 18);
        editor.SetPageBorder(RtfPageBorderSide.Right, RtfPageBorderStyle.Shadow, width: 10, colorIndex: 3);

        const string expected = @"{\rtf1\ansi\pgbrdrhead\pgbrdrfoot\pgbrdropt1\pgbrdrsnap\pgbrdrt\brdrs\brdrw12\brsp24\brdrcf1\pgbrdrb\brdrdb\brdrw18\pgbrdrr\brdrsh\brdrw10\brdrcf3{\info{\title Keep}}\pard Body\par}";
        Assert.Equal(expected, editor.ToRtf());

        RtfPageBorders borders = editor.ToReadResult().Document.PageSetup.PageBorders;
        Assert.True(borders.IncludeHeader);
        Assert.True(borders.IncludeFooter);
        Assert.True(borders.SnapToPageBorder);
        Assert.Equal(RtfPageBorderScope.FirstPageInSection, borders.Scope);
        Assert.False(borders.DisplayBehindText);
        Assert.Equal(RtfPageBorderOffset.Text, borders.OffsetFrom);
        Assert.Equal(RtfPageBorderStyle.Single, borders.Top.Style);
        Assert.Equal(RtfPageBorderStyle.Double, borders.Bottom.Style);
        Assert.Equal(RtfPageBorderStyle.Shadow, borders.Right.Style);
        Assert.True(borders.Right.Shadow);

        editor.SetPageBorderOptions();
        editor.RemovePageBorder(RtfPageBorderSide.Top);
        editor.RemovePageBorder(RtfPageBorderSide.Bottom);
        editor.RemovePageBorder(RtfPageBorderSide.Right);

        Assert.Equal(rtf, editor.ToRtf());
        Assert.False(editor.ToReadResult().Document.PageSetup.PageBorders.HasAnyValue);
    }
}
