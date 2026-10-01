using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfNativeStyleWritingTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableFallbackMarkersMaterializeCharacterStylesAndReportTheLostInheritance(bool nested) {
        RtfDocument source = RtfDocument.Create();
        RtfStyle marker = source.AddStyle(3, "Marker", RtfStyleKind.Character);
        marker.Bold = true;
        marker.TextHidden = true;
        RtfTableCell outer = source.AddTable(1, 1).Rows[0].Cells[0];
        outer.AddParagraph("Outer").SetListText(text => text.AddText("1.").StyleId = marker.Id);
        if (nested) outer.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("Inner").SetListText(text => text.AddText("2.").StyleId = marker.Id);
        var result = source.ToRtfResult();
        Assert.Equal("Paragraphs=0;Runs=" + (nested ? "2" : "1"), Assert.Single(result.Report.Diagnostics).Detail);
        RtfTable table = Assert.Single(RtfDocument.Read(result.Value).Document.Blocks.OfType<RtfTable>());
        AssertMarker(table.Rows[0].Cells[0].Paragraphs.Single(paragraph => paragraph.ToPlainText() == "Outer"));
        if (nested) AssertMarker(Assert.Single(table.Rows[0].Cells[0].Blocks.OfType<RtfTable>()).Rows[0].Cells[0].Paragraphs.Single(paragraph => paragraph.ToPlainText() == "Inner"));

        static void AssertMarker(RtfParagraph paragraph) {
            RtfRun run = Assert.Single(paragraph.ListText!.Runs);
            Assert.True(run.DirectBold);
            Assert.True(run.DirectHidden);
        }
    }

    [Fact]
    public void MaterializedOutputWritesEffectiveFormattingAndReportsLostInheritanceWithoutMutatingSource() {
        const string input = @"{\rtf1\ansi{\stylesheet{\s0 Normal;}{\s1\b\v\qc\li720 Styled;}}
\pard\s1 Inherited {\b0\v0 Visible} {\plain Plain}\par}";
        RtfDocument source = RtfDocument.Read(input).Document;
        var result = source.ToRtfResult();
        RtfConversionDiagnostic loss = Assert.Single(result.Report.Diagnostics);
        Assert.Equal("RtfNormalizationFormattingInheritanceMaterialized", loss.Code);
        Assert.Equal(RtfConversionAction.Flattened, loss.Action);
        Assert.Equal("Paragraphs=1;Runs=3", loss.Detail);
        Assert.Equal(4, loss.Count);
        Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());

        RtfParagraph original = source.Paragraphs[0];
        Assert.Null(original.DirectAlignment);
        Assert.Null(original.Runs[0].DirectBold);
        Assert.Null(original.Runs[0].DirectHidden);
        RtfParagraph reopened = RtfDocument.Read(result.Value).Document.Paragraphs[0];
        Assert.Equal(RtfTextAlignment.Center, reopened.DirectAlignment);
        Assert.Equal(720, reopened.LeftIndentTwips);
        Assert.True(reopened.Runs[0].DirectBold);
        Assert.True(reopened.Runs[0].DirectHidden);
        Assert.False(reopened.Runs.Single(run => run.Text == "Visible").DirectBold);
        Assert.False(reopened.Runs.Single(run => run.Text == "Visible").DirectHidden);
        Assert.False(reopened.Runs.Single(run => run.Text == "Plain").DirectHidden);

        source.Styles.Single(style => style.Id == 1).Bold = false;
        Assert.False(source.ResolveRunFormatting(original, original.Runs[0]).Bold);
        Assert.Null(RtfDocument.Read(source.ToRtf(new RtfWriteOptions { MaterializeStyleFormatting = false })).Document.Paragraphs[0].DirectAlignment);
    }

    [Fact]
    public void MaterializationDoesNotReportUnusedStylesOrAlreadyDirectFormatting() {
        RtfDocument source = RtfDocument.Create();
        source.AddStyle(9, "Unused").Bold = true;
        source.AddParagraph("Direct").Runs[0].Bold = true;
        source.ToRtfResult(new RtfWriteOptions { MaterializeStyleFormatting = true }).RequireNoLoss();
        RtfStyle used = source.AddStyle(1, "Used");
        used.Bold = true;
        source.Paragraphs[0].StyleId = used.Id;
        source.ToRtfResult(new RtfWriteOptions { MaterializeStyleFormatting = true }).RequireNoLoss();
    }

    [Fact]
    public void MaterializationUsesEachStoryAndTableParagraphAndTheContainingFieldParagraph() {
        RtfDocument source = RtfDocument.Create();
        RtfStyle bold = source.AddStyle(1, "Bold");
        bold.Bold = true;
        RtfStyle italic = source.AddStyle(2, "Italic");
        italic.Italic = true;
        RtfParagraph body = source.AddParagraph("Body");
        body.StyleId = bold.Id;
        body.AddField("QUOTE").Result.AddText("Field");
        source.AddHeaderFooter(RtfHeaderFooterKind.Header).AddParagraph("Header").StyleId = italic.Id;
        source.AddNote(RtfNoteKind.Footnote).AddParagraph("Note").StyleId = italic.Id;
        source.AddTable(1, 1).Rows[0].Cells[0].AddParagraph("Cell").StyleId = bold.Id;

        var output = source.ToRtfResult(new RtfWriteOptions { MaterializeStyleFormatting = true });
        Assert.Equal("Paragraphs=0;Runs=5", Assert.Single(output.Report.Diagnostics).Detail);
        RtfDocument reopened = RtfDocument.Read(output.Value).Document;
        Assert.True(reopened.Paragraphs[0].Runs[0].DirectBold);
        Assert.True(Assert.Single(reopened.Paragraphs[0].Inlines.OfType<RtfField>()).Result.Runs[0].DirectBold);
        Assert.True(reopened.HeaderFooters[0].Paragraphs[0].Runs[0].DirectItalic);
        Assert.True(reopened.Notes[0].Paragraphs[0].Runs[0].DirectItalic);
        Assert.True(Assert.Single(reopened.Blocks.OfType<RtfTable>()).Rows[0].Cells[0].Paragraphs.Single(paragraph => paragraph.Runs.Count > 0).Runs[0].DirectBold);
    }
}
