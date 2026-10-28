using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void PhysicalGeometry_RoundsAtTheNativeBoundaryAndSurvivesSaving() {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        section.Margins.Left = OfficeLength.Inches(0.5);
        section.Margins.Top = OfficeLength.Points(-12.03);
        section.Margins.HeaderDistance = OfficeLength.Points(18.03);
        section.PageSettings.Width = OfficeLength.Points(612.03);
        WordParagraph paragraph = document.AddParagraph("Physical dimensions");
        paragraph.IndentationBefore = OfficeLength.Points(-18.03);
        paragraph.LineSpacingBefore = OfficeLength.Points(6.03);
        WordTableRow row = document.AddTable(1, 1).Rows[0];
        row.MinimumHeight = OfficeLength.Points(20.03);
        row.Cells[0].MarginLeft = OfficeLength.Points(3.03);

        using var bytes = new MemoryStream(document.ToBytes());
        using (var saved = WordprocessingDocument.Open(bytes, false)) {
            Body body = saved.MainDocumentPart!.Document.Body!;
            SectionProperties properties = body.Elements<SectionProperties>().Single();
            PageMargin margins = properties.GetFirstChild<PageMargin>()!;
            Assert.Equal(720U, margins.Left!.Value);
            Assert.Equal(-241, margins.Top!.Value);
            Assert.Equal(361U, margins.Header!.Value);
            Assert.Equal(12241U, properties.GetFirstChild<PageSize>()!.Width!.Value);
            ParagraphProperties formatting = body.Elements<Paragraph>().First().ParagraphProperties!;
            Assert.Equal("-361", formatting.Indentation!.Left!.Value);
            Assert.Equal("121", formatting.SpacingBetweenLines!.Before!.Value);
            Table table = body.Elements<Table>().Single();
            Assert.Equal(401U, table.Descendants<TableRowHeight>().Single().Val!.Value);
            Assert.Equal("61", table.Descendants<LeftMargin>().Single().Width!.Value);
            Assert.Empty(new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(saved));
        }

        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        Assert.Equal(OfficeLength.Points(-18.05), reopened.Paragraphs[0].IndentationBefore);
        Assert.Equal(OfficeLength.Points(20.05), reopened.Tables[0].Rows[0].MinimumHeight);
        Assert.Equal(OfficeLength.Points(3.05), reopened.Tables[0].Rows[0].Cells[0].MarginLeft);
    }

    [Fact]
    public void PhysicalGeometry_RejectsNegativeUnsignedDimensionsBeforeRoundingOrMutation() {
        using WordDocument document = WordDocument.Create();
        WordSection section = document.Sections[0];
        WordTableRow row = document.AddTable(1, 1).Rows[0];
        OfficeLength negative = OfficeLength.Emus(-1);
        string sectionBefore = section._sectionProperties.OuterXml;
        string rowBefore = row._tableRow.OuterXml;
        Assert.Throws<ArgumentOutOfRangeException>(() => section.Margins.Left = negative);
        Assert.Throws<ArgumentOutOfRangeException>(() => section.Margins.HeaderDistance = negative);
        Assert.Throws<ArgumentOutOfRangeException>(() => section.PageSettings.Width = negative);
        Assert.Throws<ArgumentOutOfRangeException>(() => row.Height = negative);
        Assert.Throws<ArgumentOutOfRangeException>(() => row.MinimumHeight = negative);
        Assert.Equal(sectionBefore, section._sectionProperties.OuterXml);
        Assert.Equal(rowBefore, row._tableRow.OuterXml);
    }
}
