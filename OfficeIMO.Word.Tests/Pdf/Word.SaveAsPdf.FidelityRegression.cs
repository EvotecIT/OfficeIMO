using System;
using System.IO;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using M = DocumentFormat.OpenXml.Math;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_PositionedTableReservesFollowingTextSpace() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(1, 1);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Positioned";
        table._tableProperties!.TablePositionProperties = new TablePositionProperties {
            HorizontalAnchor = HorizontalAnchorValues.Margin,
            TablePositionXAlignment = HorizontalAlignmentValues.Right
        };
        document.AddParagraph("Following text");

        var result = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false });

        Assert.DoesNotContain(result.Report.Warnings, warning => warning.Code == "NativePositionedTableWrapApproximation");
        Assert.Contains("Following text", OfficeIMO.Pdf.PdfTextExtractor.ExtractAllText(result.Value.ToBytes()));
    }

    [Theory]
    [InlineData(false, WordBorderStyle.Nil, 0, false, false)]
    [InlineData(true, WordBorderStyle.Nil, 0, false, false)]
    [InlineData(false, WordBorderStyle.None, 0, false, true)]
    [InlineData(true, WordBorderStyle.None, 0, false, true)]
    [InlineData(false, WordBorderStyle.Nil, 0, true, false)]
    [InlineData(false, WordBorderStyle.Nil, 100, false, true)]
    public void SaveAsPdf_OneSidedBorderOverrideResolvesSharedTableEdge(bool horizontal, WordBorderStyle overrideStyle, int cellSpacingTwips, bool opposingDirectVisible, bool expectedStroke) {
        string name = $"PdfOneSidedBorder{horizontal}-{overrideStyle}-{cellSpacingTwips}-{opposingDirectVisible}";
        string docPath = Path.Combine(_directoryWithFiles, name + ".docx");
        string pdfPath = Path.Combine(_directoryWithFiles, name + ".pdf");
        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(horizontal ? 2 : 1, horizontal ? 1 : 2);
            table.Style = WordTableStyle.TableGrid;
            if (cellSpacingTwips > 0) table.StyleDetails!.CellSpacing = (short)cellSpacingTwips;
            table._tableProperties!.TableBorders = new TableBorders(
                new TopBorder { Val = BorderValues.Nil },
                new BottomBorder { Val = BorderValues.Nil },
                new LeftBorder { Val = BorderValues.Nil },
                new RightBorder { Val = BorderValues.Nil },
                new InsideHorizontalBorder { Val = horizontal ? BorderValues.Single : BorderValues.Nil },
                new InsideVerticalBorder { Val = horizontal ? BorderValues.Nil : BorderValues.Single });
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Left";
            (horizontal ? table.Rows[1].Cells[0] : table.Rows[0].Cells[1]).Paragraphs[0].Text = "Right";
            if (horizontal) {
                table.Rows[0].Cells[0].Borders.BottomStyle = overrideStyle;
                if (opposingDirectVisible) table.Rows[1].Cells[0].Borders.TopStyle = WordBorderStyle.Single;
            } else {
                table.Rows[0].Cells[0].Borders.RightStyle = overrideStyle;
                if (opposingDirectVisible) table.Rows[0].Cells[1].Borders.LeftStyle = WordBorderStyle.Single;
            }
            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
        }

        string text = OfficeIMO.Pdf.PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Left", text);
        Assert.Contains("Right", text);
        Assert.Equal(expectedStroke, PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath)).Contains(" RG", StringComparison.Ordinal));
    }

    [Fact]
    public void SaveAsPdf_HeadingAndTocPreserveInlineEquationAndSimpleField() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfHeadingEquationAndField.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfHeadingEquationAndField.pdf");
        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddTableOfContent();
            WordParagraph heading = document.AddParagraph("Equation ").SetStyle(WordParagraphStyles.Heading1);
            heading._paragraph.Append(new M.OfficeMath(new M.Run(new M.Text("math-token"))));
            heading._paragraph.Append(new SimpleField(new Run(new Text(" 2020") { Space = SpaceProcessingModeValues.Preserve })) {
                Instruction = " DATE \\@ \"yyyy\" "
            });
            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
        }

        string text = OfficeIMO.Pdf.PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.True(text.Split(new[] { "Equation math-token 2020" }, StringSplitOptions.None).Length >= 3, text);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void SaveAsPdf_ConditionalBorderResolvesInheritedSharedTableEdge(bool isNil, bool expectedStroke) {
        BorderValues borderStyle = isNil ? BorderValues.Nil : BorderValues.None;
        string docPath = Path.Combine(_directoryWithFiles, $"PdfConditionalSharedEdge{borderStyle}.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, $"PdfConditionalSharedEdge{borderStyle}.pdf");
        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "PdfConditionalSharedEdge";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = styleId },
                new StyleTableProperties(new TableBorders(
                    new TopBorder { Val = BorderValues.Nil },
                    new BottomBorder { Val = BorderValues.Nil },
                    new LeftBorder { Val = BorderValues.Nil },
                    new RightBorder { Val = BorderValues.Nil },
                    new InsideHorizontalBorder { Val = BorderValues.Nil },
                    new InsideVerticalBorder { Val = BorderValues.Single })),
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(new TableCellBorders(
                        new RightBorder { Val = borderStyle })))
                { Type = TableStyleOverrideValues.FirstColumn })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(1, 2);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingFirstColumn = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Conditional left";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Conditional right";
            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
        }

        string text = OfficeIMO.Pdf.PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Conditional left", text);
        Assert.Contains("Conditional right", text);
        Assert.Equal(expectedStroke, PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath)).Contains(" RG", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SaveAsPdf_ConditionalHiddenOuterBorderOverridesTableGrid(bool isNil) {
        BorderValues borderStyle = isNil ? BorderValues.Nil : BorderValues.None;
        string docPath = Path.Combine(_directoryWithFiles, $"PdfConditionalOuterEdge{borderStyle}.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, $"PdfConditionalOuterEdge{borderStyle}.pdf");
        using (WordDocument document = WordDocument.Create(docPath)) {
            const string styleId = "PdfConditionalOuterEdge";
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            styles.Append(new Style(
                new StyleName { Val = styleId },
                new StyleTableProperties(new TableBorders(
                    new TopBorder { Val = BorderValues.Single },
                    new BottomBorder { Val = BorderValues.Nil },
                    new LeftBorder { Val = BorderValues.Nil },
                    new RightBorder { Val = BorderValues.Nil })),
                new TableStyleProperties(
                    new TableStyleConditionalFormattingTableCellProperties(new TableCellBorders(
                        new TopBorder { Val = borderStyle })))
                { Type = TableStyleOverrideValues.FirstRow })
            { Type = StyleValues.Table, StyleId = styleId });

            WordTable table = document.AddTable(1, 1);
            table._tableProperties!.TableStyle = new TableStyle { Val = styleId };
            table.ConditionalFormattingFirstRow = true;
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Hidden outer edge";
            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
        }

        Assert.Contains("Hidden outer edge", OfficeIMO.Pdf.PdfTextExtractor.ExtractAllText(pdfPath));
        Assert.DoesNotContain(" RG", PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath)));
    }

    [Fact]
    public void SaveAsPdf_MergedCellSharedEdgeResolvesEachNeighborRow() {
        double[] mixed = ExportBorderLengths(WordBorderStyle.Nil, WordBorderStyle.None, "Mixed");
        double[] hidden = ExportBorderLengths(WordBorderStyle.Nil, WordBorderStyle.Nil, "Hidden");
        double[] visible = ExportBorderLengths(WordBorderStyle.None, WordBorderStyle.None, "Visible");

        Assert.Empty(hidden);
        Assert.Single(mixed);
        Assert.Single(visible);
        Assert.True(mixed[0] > 5D && mixed[0] < visible[0] - 5D);

        double[] ExportBorderLengths(WordBorderStyle upperBorder, WordBorderStyle lowerBorder, string suffix) {
            string docPath = Path.Combine(_directoryWithFiles, $"PdfMergedSharedEdge{suffix}.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, $"PdfMergedSharedEdge{suffix}.pdf");
            using (WordDocument document = WordDocument.Create(docPath)) {
                WordTable table = document.AddTable(2, 2);
                table._tableProperties!.TableBorders = new TableBorders(
                    new TopBorder { Val = BorderValues.Nil },
                    new BottomBorder { Val = BorderValues.Nil },
                    new LeftBorder { Val = BorderValues.Nil },
                    new RightBorder { Val = BorderValues.Nil },
                    new InsideHorizontalBorder { Val = BorderValues.Nil },
                    new InsideVerticalBorder { Val = BorderValues.Single });
                table.Rows[0].Height = 650;
                table.Rows[1].Height = 850;
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Merged";
                table.Rows[0].Cells[0].MergeVertically(1);
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Upper";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Lower";
                table.Rows[0].Cells[1].Borders.LeftStyle = upperBorder;
                table.Rows[1].Cells[1].Borders.LeftStyle = lowerBorder;
                document.Save();
                document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
            }

            string operators = PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath));
            return Regex.Matches(operators, @"(?<x>-?\d+(?:\.\d+)?) (?<top>-?\d+(?:\.\d+)?) m\s+\k<x> (?<bottom>-?\d+(?:\.\d+)?) l")
                .Cast<Match>()
                .Select(match => Math.Abs(
                    double.Parse(match.Groups["top"].Value, CultureInfo.InvariantCulture) -
                    double.Parse(match.Groups["bottom"].Value, CultureInfo.InvariantCulture)))
                .ToArray();
        }
    }

    [Fact]
    public void SaveAsPdf_MergedCellSharedEdgeResolvesEachNeighborColumn() {
        double[] mixed = ExportBorderLengths(WordBorderStyle.Nil, WordBorderStyle.None, "Mixed");
        double[] hidden = ExportBorderLengths(WordBorderStyle.Nil, WordBorderStyle.Nil, "Hidden");
        double[] visible = ExportBorderLengths(WordBorderStyle.None, WordBorderStyle.None, "Visible");

        Assert.Empty(hidden);
        Assert.Single(mixed);
        Assert.Single(visible);
        Assert.True(mixed[0] > 5D && mixed[0] < visible[0] - 5D);

        double[] ExportBorderLengths(WordBorderStyle leftBorder, WordBorderStyle rightBorder, string suffix) {
            string docPath = Path.Combine(_directoryWithFiles, $"PdfMergedHorizontalSharedEdge{suffix}.docx");
            string pdfPath = Path.Combine(_directoryWithFiles, $"PdfMergedHorizontalSharedEdge{suffix}.pdf");
            using (WordDocument document = WordDocument.Create(docPath)) {
                WordTable table = document.AddTable(2, 2);
                table._tableProperties!.TableBorders = new TableBorders(
                    new TopBorder { Val = BorderValues.Nil },
                    new BottomBorder { Val = BorderValues.Nil },
                    new LeftBorder { Val = BorderValues.Nil },
                    new RightBorder { Val = BorderValues.Nil },
                    new InsideHorizontalBorder { Val = BorderValues.Single },
                    new InsideVerticalBorder { Val = BorderValues.Nil });
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Merged";
                table.Rows[0].Cells[0].MergeHorizontally(1);
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Left";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Right";
                table.Rows[1].Cells[0].Borders.TopStyle = leftBorder;
                table.Rows[1].Cells[1].Borders.TopStyle = rightBorder;
                document.Save();
                document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
            }

            string operators = PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath));
            return Regex.Matches(operators, @"(?<left>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) m\s+(?<right>-?\d+(?:\.\d+)?) \k<y> l")
                .Cast<Match>()
                .Select(match => Math.Abs(
                    double.Parse(match.Groups["left"].Value, CultureInfo.InvariantCulture) -
                    double.Parse(match.Groups["right"].Value, CultureInfo.InvariantCulture)))
                .ToArray();
        }
    }

    [Fact]
    public void SaveAsPdf_SplitMergedCellBorderSegmentsStayWithinTheirPageFragment() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfSplitMergedBorderSegments.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfSplitMergedBorderSegments.pdf");
        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(2, 2);
            table._tableProperties!.TableBorders = new TableBorders(
                new TopBorder { Val = BorderValues.Nil },
                new BottomBorder { Val = BorderValues.Nil },
                new LeftBorder { Val = BorderValues.Nil },
                new RightBorder { Val = BorderValues.Nil },
                new InsideHorizontalBorder { Val = BorderValues.Single },
                new InsideVerticalBorder { Val = BorderValues.Nil });
            table.Rows[0].Cells[0].Paragraphs[0].Text = string.Join(" ", Enumerable.Repeat("Merged continuation", 100));
            table.Rows[0].Cells[0].MergeHorizontally(1);
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Left";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Right";
            table.Rows[1].Cells[0].Borders.TopStyle = WordBorderStyle.Nil;
            table.Rows[1].Cells[1].Borders.TopStyle = WordBorderStyle.None;
            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions {
                IncludePageNumbers = false,
                PageSize = new OfficeIMO.Pdf.PageSize(300, 220),
                Margins = OfficeIMO.Pdf.PageMargins.Uniform(24)
            });
        }

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages > 1);
        string operators = PdfOperatorSearchText.From(File.ReadAllBytes(pdfPath));
        Match[] horizontalLines = Regex.Matches(operators, @"(?<left>-?\d+(?:\.\d+)?) (?<y>-?\d+(?:\.\d+)?) m\s+(?<right>-?\d+(?:\.\d+)?) \k<y> l")
            .Cast<Match>().ToArray();
        Assert.NotEmpty(horizontalLines);
        Assert.All(horizontalLines, line => Assert.True(
            double.Parse(line.Groups["right"].Value, CultureInfo.InvariantCulture) >
            double.Parse(line.Groups["left"].Value, CultureInfo.InvariantCulture),
            "A split merged cell emitted a zero or reversed horizontal border."));
    }

    [Fact]
    public void SaveAsPdf_VerticallyMergedCellUsesContinuationAlignment() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfMergedContinuationAlignment.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfMergedContinuationAlignment.pdf");
        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(3, 2);
            for (int row = 0; row < 3; row++) {
                table.Rows[row].Height = 500;
                table.Rows[row].Cells[1].Paragraphs[0].Text = "Peer" + row;
            }
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Merged";
            table.Rows[0].Cells[0].MergeVertically(2);
            table.Rows[2].Cells[0].VerticalAlignment = WordTableVerticalAlignment.Center;
            document.Save();
            document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
        }

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(pdfPath);
        var words = pdf.GetPage(1).GetWords();
        double mergedY = Assert.Single(words, word => word.Text == "Merged").BoundingBox.Bottom;
        double middleY = Assert.Single(words, word => word.Text == "Peer1").BoundingBox.Bottom;
        Assert.InRange(Math.Abs(mergedY - middleY), 0D, 8D);
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, true)]
    [InlineData(2, false)]
    [InlineData(3, true)]
    [InlineData(4, false)]
    [InlineData(5, true)]
    [InlineData(6, false)]
    [InlineData(7, true)]
    [InlineData(8, true)]
    public void SaveAsPdf_TableBorderVisibility_FollowsWordOverrides(int borderMode, bool expectedStroke) {
        string docPath = Path.Combine(_directoryWithFiles, $"PdfBorderVisibility{borderMode}.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, $"PdfBorderVisibility{borderMode}.pdf");
        using (WordDocument document = WordDocument.Create(docPath)) {
            WordTable table = document.AddTable(2, 2);
            if (borderMode == 0 || borderMode == 3 || borderMode == 7) {
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
            } else if (borderMode == 7 || borderMode == 8) {
                table._tableProperties!.TableBorders = new TableBorders();
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
                DefaultTableBorders = borderMode == 3 || borderMode == 7
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
    public void SaveAsPdf_DoesNotRenderInstructionFromMixedRunWithEmptyResult() {
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfEmptyFieldResultVisibility.pdf");
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body ")._paragraph.Append(new Run(
            new FieldChar { FieldCharType = FieldCharValues.Begin },
            new Text("Hidden body instruction"),
            new FieldChar { FieldCharType = FieldCharValues.Separate },
            new FieldChar { FieldCharType = FieldCharValues.End }));
        WordParagraph cell = document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
        cell._paragraph.Append(new Run(
            new FieldChar { FieldCharType = FieldCharValues.Begin },
            new Text("Hidden cell instruction"),
            new FieldChar { FieldCharType = FieldCharValues.Separate },
            new FieldChar { FieldCharType = FieldCharValues.End }));

        document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });
        string pdfText = OfficeIMO.Pdf.PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Body", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden body instruction", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain("Hidden cell instruction", pdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_DoesNotRenderSimpleFieldInsideComplexInstruction() {
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfComplexInstructionNestedField.pdf");
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        paragraph._paragraph.Append(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new SimpleField(new Run(new Text("Hidden"))) { Instruction = " PAGE " },
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new SimpleField(new Run(new Text("Visible"))) { Instruction = " PAGE " },
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        document.SaveAsPdf(pdfPath, new WordToPdfOptions { IncludePageNumbers = false });

        string pdfText = OfficeIMO.Pdf.PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.DoesNotContain("Hidden", pdfText, StringComparison.Ordinal);
        Assert.Contains("Visible", pdfText, StringComparison.Ordinal);
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
