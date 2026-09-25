using System;
using System.IO;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(0, false)]
    [InlineData(1, true)]
    [InlineData(2, false)]
    [InlineData(3, true)]
    [InlineData(4, false)]
    [InlineData(5, true)]
    [InlineData(6, false)]
    public void SaveAsPdf_TableBorderVisibility_FollowsWordOverrides(int borderMode, bool expectedStroke) {
        string docPath = Path.Combine(_directoryWithFiles, $"PdfBorderVisibility{borderMode}.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, $"PdfBorderVisibility{borderMode}.pdf");
        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(2, 2);
            if (borderMode == 0 || borderMode == 3) {
                table._tableProperties!.TableStyle?.Remove();
            } else {
                table.Style = WordTableStyle.TableGrid;
            }
            if (borderMode == 4) {
                table._tableProperties!.TableBorders = new TableBorders(
                    new TopBorder { Val = BorderValues.Nil },
                    new BottomBorder { Val = BorderValues.Nil },
                    new LeftBorder { Val = BorderValues.Nil },
                    new RightBorder { Val = BorderValues.Nil },
                    new InsideHorizontalBorder { Val = BorderValues.Nil },
                    new InsideVerticalBorder { Val = BorderValues.Nil });
            } else if (borderMode == 5) {
                table._tableProperties!.TableBorders = new TableBorders(
                    new TopBorder { Val = BorderValues.Nil });
            }
            for (int row = 0; row < 2; row++) {
                for (int column = 0; column < 2; column++) {
                    WordTableCell cell = table.Rows[row].Cells[column];
                    cell.Paragraphs[0].Text = $"Cell {row}{column}";
                    if (borderMode == 2 || borderMode == 6) {
                        WordBorderStyle hidden = borderMode == 2 ? WordBorderStyle.Nil : WordBorderStyle.None;
                        cell.Borders.TopStyle = hidden;
                        cell.Borders.BottomStyle = hidden;
                        cell.Borders.LeftStyle = hidden;
                        cell.Borders.RightStyle = hidden;
                    }
                }
            }
            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                IncludePageNumbers = false,
                DefaultTableBorders = borderMode == 3
            });
        }

        string pdfText = OfficeIMO.Pdf.PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Cell 00", pdfText);
        Assert.Contains("Cell 11", pdfText);
        string operators = PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath));
        Assert.Equal(expectedStroke, operators.Contains(" RG", StringComparison.Ordinal));
        if (borderMode == 0) {
            Assert.DoesNotContain("0.95 0.95 0.95 rg", operators);
        }
    }

    [Fact]
    public void SaveAsPdf_PreservesSimpleFieldResultsWithinParagraphAndTableCell() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfSimpleDateFields.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfSimpleDateFields.pdf");
        using (WordDocument document = WordDocument.Create(docPath)) {
            document.BuiltinDocumentProperties.Created = new DateTime(2020, 1, 3);
            WordParagraph paragraph = document.AddParagraph("Issued ");
            paragraph._paragraph.Append(new SimpleField(new Run(new Text("stale-date") { Space = SpaceProcessingModeValues.Preserve })) {
                Instruction = " DATE \\@ \"yyyy-MM-dd\" "
            });
            paragraph.AddText(" approved");

            WordTable table = document.AddTable(1, 1);
            WordParagraph cellParagraph = table.Rows[0].Cells[0].Paragraphs[0];
            cellParagraph.Text = "Cell date ";
            cellParagraph._paragraph.Append(new SimpleField(new Run(new Text("stale-created") { Space = SpaceProcessingModeValues.Preserve })) {
                Instruction = " CREATEDATE \\@ \"yyyy-MM-dd\" "
            });
            cellParagraph.AddText(" done");

            WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions {
                CurrentDateTime = new DateTime(2020, 1, 2)
            });
            Assert.Equal(2, report.UpdatedCount);
            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
        }

        string pdfText = OfficeIMO.Pdf.PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Issued 2020-01-02 approved", pdfText);
        Assert.Contains("Cell date 2020-01-03 done", pdfText);
    }

    [Fact]
    public void SaveAsPdf_PreservesSimpleFieldResultsInListAndTocHeading() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfSimpleFieldsListToc.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfSimpleFieldsListToc.pdf");
        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddTableOfContent();
            WordParagraph heading = document.AddParagraph("Release ").SetStyle(WordParagraphStyles.Heading1);
            heading._paragraph.Append(new SimpleField(new Run(new Text("2020") { Space = SpaceProcessingModeValues.Preserve })) {
                Instruction = " DATE \\@ \"yyyy\" "
            });
            WordParagraph listItem = document.AddList(WordListStyle.Numbered).AddItem("Due ");
            listItem._paragraph.Append(new SimpleField(new Run(new Text("2020-01-02") { Space = SpaceProcessingModeValues.Preserve })) {
                Instruction = " DATE \\@ \"yyyy-MM-dd\" "
            });
            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
        }

        string pdfText = OfficeIMO.Pdf.PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.True(pdfText.Split(new[] { "Release 2020" }, StringSplitOptions.None).Length >= 3);
        Assert.Contains("Due 2020-01-02", pdfText);
    }

    [Fact]
    public void SaveAsPdf_ConditionalNilBordersSuppressInheritedGrid() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfConditionalNilBorders.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfConditionalNilBorders.pdf");
        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "PdfConditionalNilGrid";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = styleId },
                new StyleTableProperties(new TableBorders(
                    new TopBorder { Val = BorderValues.Single },
                    new BottomBorder { Val = BorderValues.Single },
                    new LeftBorder { Val = BorderValues.Single },
                    new RightBorder { Val = BorderValues.Single },
                    new InsideHorizontalBorder { Val = BorderValues.Single },
                    new InsideVerticalBorder { Val = BorderValues.Single })),
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(new TableCellBorders(
                        new TopBorder { Val = BorderValues.Nil },
                        new BottomBorder { Val = BorderValues.Nil },
                        new LeftBorder { Val = BorderValues.Nil },
                        new RightBorder { Val = BorderValues.Nil })))
                { Type = TableStyleOverrideValues.FirstRow })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingFirstRow = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Hidden conditional grid";
            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
        }

        string pdfText = OfficeIMO.Pdf.PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Hidden conditional", pdfText);
        Assert.Contains("grid", pdfText);
        Assert.DoesNotContain(" RG", PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath)));
    }

    [Fact]
    public void SaveAsPdf_DirectNilSidePreservesConfiguredCellBorderColor() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfConfiguredCellBorder.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfConfiguredCellBorder.pdf");
        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle?.Remove();
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Paragraphs[0].Text = "Configured blue border";
            cell.Borders.TopStyle = WordBorderStyle.Nil;
            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                IncludePageNumbers = false,
                PdfOptions = new OfficeIMO.Pdf.PdfOptions {
                    DefaultTableStyle = new OfficeIMO.Pdf.PdfTableStyle {
                        BorderColor = OfficeIMO.Pdf.PdfColor.Black,
                        BorderWidth = 0.5D,
                        CellBorders = new Dictionary<(int Row, int Column), OfficeIMO.Pdf.PdfCellBorder> {
                            [(0, 0)] = new OfficeIMO.Pdf.PdfCellBorder {
                                Color = OfficeIMO.Pdf.PdfColor.FromRgb(0, 0, 255),
                                Width = 2D
                            }
                        }
                    }
                }
            });
        }

        string operators = PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath));
        Assert.Contains("0 0 1 RG", operators);
        Assert.DoesNotContain("0 0 0 RG", operators);
    }
}
