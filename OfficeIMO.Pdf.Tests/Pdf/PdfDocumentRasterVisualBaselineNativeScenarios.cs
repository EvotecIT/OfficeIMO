using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentRasterVisualBaselineTests {
    private static byte[] CreateNativeWordReport() {
        string workDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.WordNativePdfRaster", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);
        string docPath = Path.Combine(workDir, "native-word-report.docx");
        string pdfPath = Path.Combine(workDir, "native-word-report.pdf");
        string logoPath = Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png");

        try {
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddHeadersAndFooters();
                WordParagraph headerLogo = document.Sections[0].Header.Default!.AddParagraph();
                headerLogo.ParagraphAlignment = W.JustificationValues.Right;
                headerLogo.AddImage(logoPath, 48, 24);

                document.AddTableOfContent();
                document.AddParagraph("Native Word PDF Visual Gate").SetStyle(WordParagraphStyles.Heading1);

                WordParagraph summary = document.AddParagraph("This document is generated as DOCX first, then rendered with the first-party OfficeIMO PDF engine.");
                summary.SetFontSize(10);

                document.AddParagraph("Native proof areas").SetStyle(WordParagraphStyles.Heading2);

                WordParagraph styled = document.AddParagraph();
                styled.AddText("Scoped runs: ");
                styled.AddText("large blue").SetFontSize(15).ColorHex = "1f4e79";
                styled.AddText(", ");
                styled.AddText("highlighted").SetHighlight(W.HighlightColorValues.Yellow);
                styled.AddText(", and restored default text.");

                WordParagraph panel = document.AddParagraph("Shaded paragraph with uniform Word borders mapped through the native PDF path.");
                panel.ShadingFillColorHex = "e6f2ff";
                panel.Borders.TopStyle = W.BorderValues.Single;
                panel.Borders.BottomStyle = W.BorderValues.Single;
                panel.Borders.LeftStyle = W.BorderValues.Single;
                panel.Borders.RightStyle = W.BorderValues.Single;
                panel.Borders.TopColorHex = "336699";
                panel.Borders.BottomColorHex = "336699";
                panel.Borders.LeftColorHex = "336699";
                panel.Borders.RightColorHex = "336699";
                panel.Borders.TopSize = 8;
                panel.Borders.BottomSize = 8;
                panel.Borders.LeftSize = 8;
                panel.Borders.RightSize = 8;

                WordList bullets = document.AddList(WordListStyle.Bulleted);
                bullets.AddItem("Native list mapping keeps markers and text aligned.");
                bullets.AddItem("The visual gate catches rhythm drift before QuestPDF removal.");

                WordList steps = document.AddCustomList();
                steps.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
                steps.AddItem("Generated TOC appears before content.");
                steps.AddItem("Tables and lists remain aligned.");

                document.AddParagraph("Approved native checkbox").AddCheckBox(true, "Native Approved", "NativeApproved");
                document.AddParagraph("Deferred native checkbox").AddCheckBox(false, "Native Deferred", "NativeDeferred");

                document.AddParagraph("Native evidence table").SetStyle(WordParagraphStyles.Heading2);

                WordTable table = document.AddTable(4, 3);
                table.Style = WordTableStyle.GridTable1LightAccent1;
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Area";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Native status";
                table.Rows[0].Cells[2].Paragraphs[0].Text = "Evidence";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Runs";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Improving";
                table.Rows[1].Cells[2].Paragraphs[0].Text = "color, size, highlight";
                table.Rows[2].Cells[0].Paragraphs[0].Text = "Tables";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Partial";
                table.Rows[2].Cells[2].Paragraphs[0].Text = "style and borders";
                table.Rows[3].Cells[0].Paragraphs[0].Text = "Forms";
                table.Rows[3].Cells[1].Paragraphs[0].Text = "Improving";
                table.Rows[3].Cells[2].Paragraphs[0].Text = "cell checkbox";
                table.Rows[3].Cells[2].Paragraphs[0].AddCheckBox(true, "Native Table Approved", "NativeTableApproved");
                table.RepeatHeaderRowAtTheTopOfEachPage = true;

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new OfficeIMO.Pdf.PageSize(612, 792),
                    Margins = OfficeIMO.Pdf.PageMargins.Uniform(36)
                });
            }

            return File.ReadAllBytes(pdfPath);
        } finally {
            TryDeleteDirectory(workDir);
        }
    }

    private static byte[] CreateNativeWordDailyLayout() {
        string workDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.WordNativePdfDailyLayout", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);
        string docPath = Path.Combine(workDir, "native-word-daily-layout.docx");
        string pdfPath = Path.Combine(workDir, "native-word-daily-layout.pdf");
        string logoPath = Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png");

        try {
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.Settings.FontFamily = "Calibri";
                document.Background.SetColorHex("EAF4FF");
                WordSection section = document.Sections[0];
                section.Margins.LeftCentimeters = 1.6;
                section.Margins.RightCentimeters = 1.4;
                section.Margins.TopCentimeters = 1.4;
                section.Margins.BottomCentimeters = 1.6;
                section.ColumnCount = 2;
                section.ColumnsSpace = 540;
                section.HasColumnSeparator = true;

                document.AddHeadersAndFooters();
                WordParagraph header = document.Sections[0].Header.Default!.AddParagraph();
                header.ParagraphAlignment = W.JustificationValues.Right;
                header.AddText("Daily layout gate");
                header.AddImage(logoPath, 42, 20);
                document.Sections[0].Footer.Default!.AddParagraph("OfficeIMO native Word layout proof");

                document.AddTableOfContent();
                document.AddParagraph("Daily Word Layout Gate").SetStyle(WordParagraphStyles.Heading1);

                WordParagraph intro = document.AddParagraph("This Word-origin fixture combines margins, section columns, fonts, colors, links, images, lists, a TOC, and a table on one page.");
                intro.SetFontSize(9);

                WordParagraph leftHeading = document.AddParagraph("Column narrative");
                leftHeading.SetStyle(WordParagraphStyles.Heading2);
                leftHeading.SetColorHex("#1f4e79");

                WordParagraph rich = document.AddParagraph();
                rich.AddText("Styled runs: ");
                rich.AddText("Calibri default").SetFontFamily("Calibri");
                rich.AddText(", ");
                rich.AddText("Courier note").SetFontFamily("Courier New").SetFontSize(8);
                rich.AddText(", ");
                rich.AddText("blue emphasis").SetFontSize(11).ColorHex = "1f4e79";
                rich.AddText(", and ");
                rich.AddHyperLink("external link", new Uri("https://evotec.xyz/native-daily-layout"), addStyle: true, tooltip: "Native daily layout link");
                rich.AddText(".");

                WordParagraph shaded = document.AddParagraph("Shaded paragraph and side borders protect report-style callouts.");
                shaded.ShadingFillColorHex = "e6f2ff";
                shaded.Borders.LeftStyle = W.BorderValues.Single;
                shaded.Borders.LeftColorHex = "1f4e79";
                shaded.Borders.LeftSize = 12;
                shaded.Borders.BottomStyle = W.BorderValues.Single;
                shaded.Borders.BottomColorHex = "b7c9d9";
                shaded.Borders.BottomSize = 4;

                WordList bullets = document.AddList(WordListStyle.Bulleted);
                bullets.AddItem("Bullet markers stay inside the first Word column.");
                bullets.AddItem("Wrapped list text keeps readable spacing.");

                WordList steps = document.AddCustomList();
                steps.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
                steps.AddItem("Collect daily document patterns.");
                steps.AddItem("Render through OfficeIMO.Pdf.");

                document.AddParagraph().AddImage(logoPath, 54, 26);
                WordParagraph columnHandoff = document.AddParagraph("Inline column break keeps this text in the first column.");
                columnHandoff.AddBreak(W.BreakValues.Column);
                columnHandoff.AddText("Right column starts here.");

                WordParagraph rightHeading = document.AddParagraph("Column evidence");
                rightHeading.SetStyle(WordParagraphStyles.Heading2);
                rightHeading.SetColorHex("#2f6f3e");

                WordTable table = document.AddTable(4, 3);
                table.Style = WordTableStyle.GridTable1LightAccent1;
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Area";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Mapped";
                table.Rows[0].Cells[2].Paragraphs[0].Text = "Visual";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Margins";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Yes";
                table.Rows[1].Cells[2].Paragraphs[0].Text = "left edge";
                table.Rows[2].Cells[0].Paragraphs[0].Text = "Columns";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Yes";
                table.Rows[2].Cells[2].Paragraphs[0].Text = "separator";
                table.Rows[3].Cells[0].Paragraphs[0].Text = "Links";
                table.Rows[3].Cells[1].Paragraphs[0].AddHyperLink("Yes", new Uri("https://officeimo.net/"), addStyle: true, tooltip: "Native table link");
                table.Rows[3].Cells[2].Paragraphs[0].Text = "annotation";

                WordParagraph closing = document.AddParagraph("Footer, header image, TOC link, section separator, and table borders should all survive the native PDF path.");
                closing.SetFontSize(8);

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new PageSize(612, 792)
                });
            }

            return File.ReadAllBytes(pdfPath);
        } finally {
            TryDeleteDirectory(workDir);
        }
    }

    private static byte[] CreateNativeWordTableCellPictureControl() {
        string workDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.WordNativePdfRaster", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);
        string docPath = Path.Combine(workDir, "native-word-table-cell-picture-control.docx");
        string pdfPath = Path.Combine(workDir, "native-word-table-cell-picture-control.pdf");
        string logoPath = Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png");

        try {
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Native table-cell picture control").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("The logo below starts as a Word picture content control inside a table cell and renders through OfficeIMO.Pdf.");

                WordTable table = document.AddTable(2, 2, WordTableStyle.TableGrid);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Control";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Rendered evidence";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "Picture content control";
                WordParagraph pictureCell = table.Rows[1].Cells[1].Paragraphs[0];
                pictureCell.Text = "Table-cell logo";
                pictureCell.ParagraphAlignment = W.JustificationValues.Center;
                pictureCell.AddPictureControl(logoPath, 72, 36, "Table Cell Logo", "TableCellLogo");

                document.Save();
                document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                    IncludePageNumbers = false,
                    PageSize = new PageSize(612, 360),
                    Margins = PageMargins.Uniform(36)
                });
            }

            return File.ReadAllBytes(pdfPath);
        } finally {
            TryDeleteDirectory(workDir);
        }
    }

    private static byte[] CreateNativeExcelDailyWorkbook() {
        string workDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.ExcelNativePdfRaster", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);
        string workbookPath = Path.Combine(workDir, "native-excel-daily-workbook.xlsx");
        string pdfPath = Path.Combine(workDir, "native-excel-daily-workbook.pdf");
        byte[] logoBytes = File.ReadAllBytes(Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png"));

        try {
            using (ExcelDocument document = ExcelDocument.Create(workbookPath)) {
                ExcelSheet summary = document.AddWorksheet("Summary");
                summary.CellAt(1, 1).SetValue("Daily Excel PDF Gate").SetBold().SetFontColor("1F4E79").SetFillColor("DDEEFF");
                summary.MergeRange("A1:B1");
                summary.CellAt(2, 1).SetValue("Metric").SetBold().SetFillColor("E6F2FF");
                summary.CellAt(2, 2).SetValue("Value").SetBold().SetFillColor("E6F2FF");
                summary.CellAt(2, 3).SetValue("Status").SetBold().SetFillColor("E6F2FF");
                summary.Cell(3, 1, "Revenue");
                summary.CellAt(3, 2).SetValue(12345.6).Currency(2, System.Globalization.CultureInfo.GetCultureInfo("en-US"));
                summary.CellAt(3, 3).SetValue("On track").SetFontColor("2F6F3E");
                summary.Cell(4, 1, "Margin");
                summary.CellAt(4, 2).SetValue(0.257).Percent(1);
                summary.SetHyperlink(4, 3, "https://officeimo.net/excel-pdf", display: "External Link");
                summary.Cell(5, 1, "Drilldown");
                summary.SetInternalLink(5, 2, "Details!A1", display: "Open Details");
                summary.Cell(6, 1, "Visual image");
                summary.Cell(6, 3, "logo above grid");
                summary.AddImage(6, 2, logoBytes, "image/png", widthPixels: 64, heightPixels: 28, name: "Summary Logo", altText: "Summary visual proof");
                summary.Cell(8, 1, "HiddenRowValue");
                summary.Cell(10, 1, "Month");
                summary.Cell(10, 2, "Actual");
                summary.Cell(10, 3, "Target");
                summary.Cell(11, 1, "Jan");
                summary.Cell(11, 2, 12);
                summary.Cell(11, 3, 10);
                summary.Cell(12, 1, "Feb");
                summary.Cell(12, 2, 18);
                summary.Cell(12, 3, 16);
                summary.Cell(13, 1, "Mar");
                summary.Cell(13, 2, 24);
                summary.Cell(13, 3, 20);
                summary.AddChartFromRange("A10:C13", row: 7, column: 1, widthPixels: 300, heightPixels: 140, type: ExcelChartType.ColumnClustered, title: "Revenue Chart");
                summary.SetColumnWidth(1, 16);
                summary.SetColumnWidth(2, 18);
                summary.SetColumnWidth(3, 28);
                summary.SetRowHeight(1, 28);
                summary.SetRowHeight(6, 32);
                summary.SetColumnHidden(4, true);
                summary.Cell(2, 4, "HiddenColumnValue");
                summary.SetRowHidden(8, true);
                summary.CellBorder(2, 1, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, "445566");
                summary.CellBorder(2, 2, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, "445566");
                summary.CellBorder(2, 3, DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues.Thin, "445566");
                summary.CellAlign(3, 2, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Right);
                summary.CellAlign(4, 2, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Right);
                summary.SetHeaderFooter(
                    headerLeft: "Daily Excel",
                    headerCenter: "Workbook &A",
                    headerRight: "Page &P of &N",
                    footerLeft: "OfficeIMO Excel PDF",
                    footerRight: "Visual baseline");
                summary.SetHeaderImage(HeaderFooterPosition.Left, logoBytes, "image/png", widthPoints: 28, heightPoints: 14);
                summary.SetOrientation(ExcelPageOrientation.Landscape);
                summary.SetMargins(left: 0.35, right: 0.35, top: 0.55, bottom: 0.55);

                ExcelSheet details = document.AddWorksheet("Details");
                details.CellAt(1, 1).SetValue("Details Target").SetBold().SetFillColor("E2F0D9");
                details.Cell(2, 1, "Owner");
                details.Cell(2, 2, "OfficeIMO");
                details.Cell(3, 1, "Status");
                details.Cell(3, 2, "Linked sheet destination");
                details.SetColumnWidth(1, 16);
                details.SetColumnWidth(2, 30);

                document.SetPrintArea(summary, "A1:C7");
                document.SetPrintTitles(summary, firstRow: 2, lastRow: 2, firstCol: null, lastCol: null);
                document.SetPrintArea(details, "A1:B3");
                document.Save();

                document.SaveAsPdf(pdfPath, new ExcelPdfSaveOptions {
                    IncludeSheetHeadings = true,
                    HeaderRowCount = 2,
                    PageSize = new PageSize(792, 612),
                    Margins = PageMargins.FromInches(0.35, 0.55, 0.35, 0.55)
                });
            }

            return File.ReadAllBytes(pdfPath);
        } finally {
            TryDeleteDirectory(workDir);
        }
    }

    private static byte[] CreateMarkdownTechnicalDocument() {
        string workDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.MarkdownPdfRaster", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(workDir);
        string markdownPath = Path.Combine(workDir, "markdown-technical-document.md");
        string logoPath = Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png");
        string localLogoPath = Path.Combine(workDir, "EvotecLogo.png");

        try {
            if (File.Exists(logoPath)) {
                File.Copy(logoPath, localLogoPath, overwrite: true);
            } else {
                File.WriteAllBytes(localLogoPath, CreateFallbackLogo());
            }

            File.WriteAllText(markdownPath, """
---
title: Markdown PDF Visual Gate
author: OfficeIMO
tags: [pdf, markdown, native]
description: Dependency-free Markdown export through OfficeIMO.Pdf
---
# Markdown PDF Visual Gate

This fixture proves that `OfficeIMO.Markdown.Pdf` can turn a practical Markdown document into a polished first-party PDF without adding another rendering dependency.

![OfficeIMO logo](EvotecLogo.png){width=104 height=36}
_Figure 1. Relative local image resolved from the Markdown file directory._

> The adapter keeps Markdown semantics thin and routes visual work through reusable PDF primitives.

## Export coverage

- [x] Headings become PDF outlines and named destinations.
- [x] Rich inline text keeps **bold**, _italic_, `code`, and [links](https://officeimo.net/).
- [x] Tables, quotes, code blocks, metadata, and local images share the core PDF engine.

| Feature | Mapping | Status |
| --- | --- | --- |
| Metadata | Front matter to PDF info dictionary | Native |
| Lists | Rich PDF list items | Native |
| Tables | First-party PDF table cells | Native |
| Images | Local JPEG/PNG image blocks | Native |

```csharp
OfficeIMO.Markdown.MarkdownDoc.Load("README.md").SaveAsPdf("README.pdf");
```
""", new UTF8Encoding(false));

            var options = new MarkdownPdfSaveOptions {
                DefaultImageWidth = 104,
                DefaultImageHeight = 36,
                ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreateTrustedHost(),
                BaseDirectory = workDir
            };

            byte[] pdf = OfficeIMO.Markdown.MarkdownDoc.Load(markdownPath).ToPdf(options);
            if (options.Warnings.Count != 0) {
                throw new InvalidOperationException("Markdown raster fixture produced export warnings: " + string.Join("; ", options.Warnings.Select(warning => warning.Code + ":" + warning.Source)));
            }

            return pdf;
        } finally {
            TryDeleteDirectory(workDir);
        }
    }

    private static byte[] CreateMarkdownThemeGallery(OfficeVisualThemeKind themeKind) {
        string markdown = CreateMarkdownThemeGallerySource(themeKind);
        var options = new MarkdownPdfSaveOptions {
            Style = MarkdownPdfStyle.Create(themeKind)
        };

        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(markdown).ToPdf(options);
        if (options.Warnings.Count != 0) {
            throw new InvalidOperationException("Markdown theme gallery fixture produced export warnings: " + string.Join("; ", options.Warnings.Select(warning => warning.Code + ":" + warning.Source)));
        }

        return pdf;
    }

    private static string CreateMarkdownThemeGallerySource(OfficeVisualThemeKind themeKind) {
        string themeName = themeKind.ToString();
        return """
# Markdown Theme Gallery

This page renders one first-party visual profile for headings, lists, tables, code, quotes, callouts, and TOC chrome.

[TOC min=2 max=2 layout=panel title="Contents" requiretoplevel=false]

> [!TIP] Theme profile: THEME_NAME
> Markdown stays semantic while `OfficeIMO.Pdf` owns visual layout.

## Evidence

- [x] Headings create PDF hierarchy.
- [x] Lists keep readable spacing.
- [x] Panels and tables are styled by the selected profile.
- [ ] Wrapped checklist rows keep the checkbox optically aligned with the first text line.

| Surface | Expected visual signal |
| --- | --- |
| Table header | Theme-specific fill and text color |
| Code block | Monospace panel with controlled spacing |

> Quotes should read as supporting narrative.

```csharp
var pdf = markdown.ToPdf();
```
""".Replace("THEME_NAME", themeName);
    }

    private static double FindWordStartX(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.X;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }
}
