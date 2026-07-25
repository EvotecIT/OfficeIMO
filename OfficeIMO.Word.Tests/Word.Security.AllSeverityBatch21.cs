using System.Reflection;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using UglyToad.PdfPig;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void ListInfoRetainsOriginalConstructorSignature() {
        ConstructorInfo? constructor = typeof(DocumentTraversal.ListInfo).GetConstructor(new[] {
            typeof(int),
            typeof(bool),
            typeof(int),
            typeof(NumberFormatValues?),
            typeof(string),
            typeof(int?),
            typeof(int?)
        });

        Assert.NotNull(constructor);
    }

    [Fact]
    public void ContinuousSectionWithExplicitRestartStartsNewPdfNumberingGroup() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeContinuousSectionRestart.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeContinuousSectionRestart.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.Sections[0].AddPageNumbering(1, NumberFormatValues.Decimal);
            document.AddParagraph("BeforeRestart");
            WordSection second = document.AddSection(SectionMarkValues.Continuous);
            second.AddPageNumbering(1, NumberFormatValues.Decimal);
            document.AddParagraph("AfterRestart");
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = true,
                PageSize = new OfficeIMO.Pdf.PageSize(400, 300),
                Margins = OfficeIMO.Pdf.PageMargins.Uniform(50)
            });
        }

        using PdfDocument pdf = PdfDocument.Open(pdfPath);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("BeforeRestart", pdf.GetPage(1).Text);
        Assert.Contains("AfterRestart", pdf.GetPage(2).Text);
    }

    [Fact]
    public void TableStyleFillPrecedenceIsDirectThenConditionalThenBase() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleFillPrecedence.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTableStyleFillPrecedence.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "NativeTableStyleFillPrecedence";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = "Native Table Style Fill Precedence" },
                new StyleTableProperties(
                    new Shading { Val = ShadingPatternValues.Clear, Fill = "D9EAD3" }),
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(
                        new Shading { Val = ShadingPatternValues.Clear, Fill = "CFE2F3" }))
                { Type = TableStyleOverrideValues.FirstRow })
            {
                Type = StyleValues.Table,
                StyleId = styleId,
                CustomStyle = true
            });

            WordTable table = document.AddTable(2, 2);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingFirstRow = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "ConditionalFill";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "DirectFill";
            table.Rows[0].Cells[1].ShadingFillColorHex = "F4CCCC";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "BaseFill";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "BasePeer";

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false,
                PageSize = new OfficeIMO.Pdf.PageSize(420, 260),
                Margins = OfficeIMO.Pdf.PageMargins.Uniform(40)
            });
        }

        string content = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.812 0.886 0.953 rg", content);
        Assert.Contains("0.957 0.8 0.8 rg", content);
        Assert.Contains("0.851 0.918 0.827 rg", content);
    }
}
