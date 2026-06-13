using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_CoverPage_And_Advances_Toc_PageNumbers() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageToc.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageToc.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlock("Native cover title", "Native cover subtitle"));
            document.AddTableOfContent();
            document.AddPageBreak();
            document.AddParagraph("Native heading after cover").SetStyle(WordParagraphStyles.Heading1);
            document.AddParagraph("Native body after cover");

            object entries = BuildNativeTableOfContentsEntries(document);
            object entry = ((System.Collections.IEnumerable)entries)
                .Cast<object>()
                .Single(item => string.Equals((string)item.GetType().GetProperty("Text")!.GetValue(item)!, "Native heading after cover", StringComparison.Ordinal));
            int pageNumber = (int)entry.GetType().GetProperty("PageNumber")!.GetValue(entry)!;
            Assert.Equal(3, pageNumber);

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 3);
        string allText = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("Native cover title", allText);
        Assert.Contains("Table of Contents", allText);
        Assert.Contains("Native heading after cover", allText);
        Assert.True(allText.IndexOf("Native cover title", StringComparison.Ordinal) < allText.IndexOf("Table of Contents", StringComparison.Ordinal));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Toc_Entries_Use_Content_Right_Edge_Tab_Stop() {
        object level1Style = CreateNativeTableOfContentsEntryStyle(relativeLevel: 0, contentWidth: 468D);
        object level2Style = CreateNativeTableOfContentsEntryStyle(relativeLevel: 1, contentWidth: 468D);
        object level3Style = CreateNativeTableOfContentsEntryStyle(relativeLevel: 2, contentWidth: 468D);

        Assert.Equal(0D, GetPdfParagraphStyleDouble(level1Style, "LeftIndent"));
        Assert.Equal(468D, GetPdfParagraphStyleDouble(level1Style, "DefaultTabStopWidth"));
        Assert.Equal(22D, GetPdfParagraphStyleDouble(level2Style, "LeftIndent"));
        Assert.Equal(446D, GetPdfParagraphStyleDouble(level2Style, "DefaultTabStopWidth"));
        Assert.Equal(44D, GetPdfParagraphStyleDouble(level3Style, "LeftIndent"));
        Assert.Equal(424D, GetPdfParagraphStyleDouble(level3Style, "DefaultTabStopWidth"));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_CoverPage_Property_ContentControls() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageProperties.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageProperties.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.BuiltinDocumentProperties.Title = "Native property title";
            document.ApplicationProperties.Company = "Native property company";
            document._document.Body!.Append(CreateNativeCoverPageBlock(
                CreateNativeCoverPagePropertyBlock("Title", "[Document title]"),
                CreateNativeCoverPagePropertyBlock("Company", "[company name]")));
            document.AddParagraph("Native property body");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Native property title", text);
        Assert.Contains("Native property company", text);
        Assert.Contains("Native property body", text);
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyContentControlUnsupported");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_CoverPage_Inline_Property_ContentControls() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageInlineProperties.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageInlineProperties.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.BuiltinDocumentProperties.Title = "Native inline title";
            document.ApplicationProperties.Company = "Native inline company";
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                new Paragraph(
                    CreateNativeCoverPagePropertyRun("Title", "[Document title]"),
                    new Run(new Text(" - ")),
                    CreateNativeCoverPagePropertyRun("Company", "[company name]"))));
            document.AddParagraph("Native inline property body");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Native inline title", text);
        Assert.Contains("Native inline company", text);
        Assert.Contains("Native inline property body", text);
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyContentControlUnsupported");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Resolves_CoverPage_TextBox_Property_Placeholders() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageTextBoxProperties.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageTextBoxProperties.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.BuiltinDocumentProperties.Title = "Native textbox title";
            document.BuiltinDocumentProperties.Created = new DateTime(2026, 6, 13);
            document.ApplicationProperties.Company = "Native textbox company";
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlTextBoxParagraph(
                    new Paragraph(
                        CreateNativeCoverPagePropertyRun("Date", "[Date]"),
                        new Run(new Text(" ")),
                        CreateNativeCoverPagePropertyRun("Company", "[company name]"),
                        new Run(new Text(" ")),
                        CreateNativeCoverPagePropertyRun("Title", "[Document title]")))));
            document.AddParagraph("Native textbox property body");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("2026-06-13", text);
        Assert.Contains("Native textbox title", text);
        Assert.Contains("Native textbox company", text);
        Assert.Contains("Native textbox property body", text);
        Assert.DoesNotContain("[Document title]", text);
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyContentControlUnsupported");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Resolves_Vml_TextBox_Alias_And_Caps() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageTextBoxAliasCaps.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageTextBoxAliasCaps.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.ApplicationProperties.Company = "Native mixed case company";
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlTextBoxParagraph(
                    new Paragraph(
                        CreateNativeCoverPagePropertyRun(
                            "Company",
                            "[bound company placeholder]",
                            new RunProperties(new Caps()))))));
            document.AddParagraph("Native alias caps body");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("NATIVE MIXED CASE COMPANY", text);
        Assert.Contains("Native alias caps body", text);
        Assert.DoesNotContain("[bound company placeholder]", text);
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyContentControlUnsupported");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Vml_CoverPage_Drawing_On_First_Page() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlDrawing.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlDrawing.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.BuiltinDocumentProperties.Title = "Native VML cover title";
            document.BuiltinDocumentProperties.Created = new DateTime(2026, 6, 13);
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlCoverDrawingParagraph()));
            document.AddParagraph("Native VML cover body");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("Native VML cover", pdf.GetPage(1).Text);
        Assert.Contains("title", pdf.GetPage(1).Text);
        Assert.Contains("2026-06-13", pdf.GetPage(1).Text);
        Assert.Contains("Native VML cover body", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains(" re", pageContent);
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyContentControlUnsupported");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Nested_Vml_CoverPage_CoordOrigin_Freeform() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlCoordOrigin.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlCoordOrigin.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlCoordOriginDrawingParagraph()));
            document.AddParagraph("After nested VML cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After nested VML cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.8 0.333 0 rg", pageContent, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Maps_Text_Watermark_To_Pdf_Watermark() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeTextWatermark.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeTextWatermark.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            WordWatermark watermark = header.AddWatermark(WordWatermarkStyle.Text, "CONFIDENTIAL");
            watermark.FontFamily = "Arial";
            watermark.ColorHex = "silver";
            watermark.Opacity = 0.35;
            document.AddParagraph("Native watermark body text");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("/GS", pageContent);
        Assert.Contains("/FW", pageContent);

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("\nC", text);
        Assert.Contains("Native watermark body text", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Normalizes_Direct_Heading_Text_LineBreaks() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeDirectHeadingLineBreak.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeDirectHeadingLineBreak.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(new Paragraph(
                new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" }),
                new Run(new Text("\nNative direct heading") { Space = SpaceProcessingModeValues.Preserve })));
            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Native direct heading", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Normalizes_HeaderFooter_Table_Cell_LineBreaks() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterCellLineBreak.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeHeaderFooterCellLineBreak.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordTable headerTable = RequireSectionHeader(document, 0, HeaderFooterValues.Default).AddTable(1, 1, WordTableStyle.TableNormal);
            headerTable.Rows[0].Cells[0].Paragraphs[0].Text = "Native header first";
            headerTable.Rows[0].Cells[0].AddParagraph("Native header second");
            document.AddParagraph("Native header newline body");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Native header first Native header second", text);
        Assert.Contains("Native header newline body", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Word_Charts_Through_Shared_Renderer() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChart.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChart.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddParagraph("Chart section").SetStyle(WordParagraphStyles.Heading1);
            WordChart chart = document.AddChart("Word PDF Chart", false, 360, 220);
            chart.AddPie("Passed", 3);
            chart.AddPie("Failed", 1);
            document.AddParagraph("After chart");

            object snapshot = CreateNativeWordChartSnapshot(chart);
            Assert.Equal(OfficeChartKind.Pie, snapshot.GetType().GetProperty("ChartKind")!.GetValue(snapshot));
            object data = snapshot.GetType().GetProperty("Data")!.GetValue(snapshot)!;
            var categories = ((System.Collections.IEnumerable)data.GetType().GetProperty("Categories")!.GetValue(data)!).Cast<string>().ToList();
            Assert.Equal(new[] { "Passed", "Failed" }, categories);

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After chart", text);

        string rawPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.122 0.306 0.475 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.184 0.435 0.243 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Word_Pie_DataLabels() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordPieDataLabels.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordPieDataLabels.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Pie Labels", false, 360, 220);
            chart.AddPie("Passed", 1);
            chart.AddPie("Failed", 0);
            chart.AddPie("Skipped", 0);
            document.AddParagraph("After pie labels");

            object snapshot = CreateNativeWordChartSnapshot(chart);
            object layout = snapshot.GetType().GetProperty("Layout")!.GetValue(snapshot)!;
            Assert.True((bool)layout.GetType().GetProperty("ShowDataLabels")!.GetValue(layout)!);
            Assert.True((bool)layout.GetType().GetProperty("ShowDataLabelValues")!.GetValue(layout)!);
            Assert.True((bool)layout.GetType().GetProperty("ShowDataLabelPercentages")!.GetValue(layout)!);

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("1; 100%", text);
        Assert.Contains("0; 0%", text);
        Assert.Contains("After pie labels", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_Embedded_Fallback_For_Word_Symbol_Text() {
        if (!PdfEmbeddedFontFamily.TryFromSystem("Segoe UI Symbol", out _) &&
            !PdfEmbeddedFontFamily.TryFromSystem("DejaVu Sans", out _)) {
            return;
        }

        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordSymbolFallback.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordSymbolFallback.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.Settings.FontFamily = "Calibri";
            WordParagraph paragraph = document.AddParagraph();
            paragraph.AddText("✓ Native symbol text").SetBold();

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.True(
            text.Contains("✓ Native symbol text", StringComparison.Ordinal) ||
            text.Contains("✓Native symbol text", StringComparison.Ordinal),
            text);
    }

    private static SdtBlock CreateNativeCoverPageBlock(params string[] lines) {
        var block = new SdtBlock(
            new SdtProperties(
                new SdtContentDocPartObject(
                    new DocPartGallery { Val = "Cover Pages" },
                    new DocPartUnique())),
            new SdtContentBlock());
        SdtContentBlock content = block.GetFirstChild<SdtContentBlock>()!;
        foreach (string line in lines) {
            content.Append(new Paragraph(new Run(new Text(line) { Space = SpaceProcessingModeValues.Preserve })));
        }

        return block;
    }

    private static object CreateNativeWordChartSnapshot(WordChart chart) {
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("TryCreateNativeWordChartSnapshot", BindingFlags.NonPublic | BindingFlags.Static)!;
        object?[] arguments = { chart, null, null };
        bool result = (bool)method.Invoke(null, arguments)!;
        Assert.True(result, (string?)arguments[2]);
        return arguments[1]!;
    }

    private static SdtBlock CreateNativeCoverPageBlockWithChildren(params OpenXmlElement[] children) {
        var block = new SdtBlock(
            new SdtProperties(
                new SdtContentDocPartObject(
                    new DocPartGallery { Val = "Cover Pages" },
                    new DocPartUnique())),
            new SdtContentBlock());
        SdtContentBlock content = block.GetFirstChild<SdtContentBlock>()!;
        foreach (OpenXmlElement child in children) {
            content.Append(child);
        }

        return block;
    }

    private static SdtBlock CreateNativeCoverPageBlock(params SdtBlock[] blocks) {
        var block = new SdtBlock(
            new SdtProperties(
                new SdtContentDocPartObject(
                    new DocPartGallery { Val = "Cover Pages" },
                    new DocPartUnique())),
            new SdtContentBlock());
        SdtContentBlock content = block.GetFirstChild<SdtContentBlock>()!;
        foreach (SdtBlock childBlock in blocks) {
            content.Append(childBlock);
        }

        return block;
    }

    private static SdtBlock CreateNativeCoverPagePropertyBlock(string alias, string placeholder) {
        return new SdtBlock(
            new SdtProperties(
                new SdtAlias { Val = alias },
                new ShowingPlaceholder()),
            new SdtContentBlock(
                new Paragraph(new Run(new Text(placeholder) { Space = SpaceProcessingModeValues.Preserve }))));
    }

    private static SdtRun CreateNativeCoverPagePropertyRun(string alias, string placeholder, RunProperties? runProperties = null) {
        var run = new Run(new Text(placeholder) { Space = SpaceProcessingModeValues.Preserve });
        if (runProperties != null) {
            run.PrependChild((RunProperties)runProperties.CloneNode(true));
        }

        return new SdtRun(
            new SdtProperties(
                new SdtAlias { Val = alias },
                new ShowingPlaceholder()),
            new SdtContentRun(
                run));
    }

    private static Paragraph CreateNativeVmlTextBoxParagraph(params OpenXmlElement[] children) {
        var textBox = new DocumentFormat.OpenXml.Vml.TextBox(
            new TextBoxContent(children));
        var shape = new DocumentFormat.OpenXml.Vml.Shape(textBox) {
            Id = "NativeCoverTextBox",
            Type = "#_x0000_t202",
            Style = "position:absolute;left:72pt;top:72pt;width:420pt;height:108pt"
        };

        return new Paragraph(new Run(new Picture(shape)));
    }

    private static Paragraph CreateNativeVmlCoverDrawingParagraph() {
        var leftBand = new DocumentFormat.OpenXml.Vml.Rectangle {
            Id = "NativeCoverLeftBand",
            Style = "position:absolute;left:24;top:24;width:18;height:720",
            FillColor = "#44546a"
        };

        var dateTextBox = new DocumentFormat.OpenXml.Vml.TextBox(
            new TextBoxContent(
                new Paragraph(
                    new ParagraphProperties(new Justification { Val = JustificationValues.Right }),
                    new Run(
                        new RunProperties(new FontSize { Val = "28" }, new Color { Val = "FFFFFF" }),
                        new Text("[Date]") { Space = SpaceProcessingModeValues.Preserve }))));
        var dateRibbon = new DocumentFormat.OpenXml.Vml.Shape(dateTextBox) {
            Id = "NativeCoverDateRibbon",
            Type = "#_x0000_t15",
            Style = "position:absolute;left:42;top:140;width:168;height:36;v-text-anchor:middle",
            FillColor = "#4472c4"
        };

        var titleTextBox = new DocumentFormat.OpenXml.Vml.TextBox(
            new TextBoxContent(
                new Paragraph(
                    new Run(
                        new RunProperties(new FontSize { Val = "52" }, new Color { Val = "262626" }),
                        new Text("[Document title]") { Space = SpaceProcessingModeValues.Preserve }))));
        var titleShape = new DocumentFormat.OpenXml.Vml.Shape(titleTextBox) {
            Id = "NativeCoverTitleTextBox",
            Type = "#_x0000_t202",
            Style = "position:absolute;mso-left-percent:430;mso-top-percent:230;width:240pt;height:80pt;v-text-anchor:top"
        };

        var group = new DocumentFormat.OpenXml.Vml.Group(leftBand, dateRibbon) {
            Id = "NativeCoverDecor",
            Style = "position:absolute;margin-left:0;margin-top:0;width:240pt;height:760pt",
            CoordinateSize = "240,760"
        };

        return new Paragraph(new Run(new Picture(group, titleShape)));
    }

    private static Paragraph CreateNativeVmlCoordOriginDrawingParagraph() {
        var freeform = new DocumentFormat.OpenXml.Vml.Shape {
            Id = "NativeCoverCoordOriginFreeform",
            Style = "position:absolute;left:1200;top:2200;width:200;height:180;visibility:visible",
            FillColor = "#CC5500"
        };
        freeform.SetAttribute(new OpenXmlAttribute("coordsize", string.Empty, "100,100"));
        freeform.SetAttribute(new OpenXmlAttribute("path", string.Empty, "m0,0l100,0,100,100,0,100xe"));

        var nested = new DocumentFormat.OpenXml.Vml.Group(freeform) {
            Id = "NativeCoverNestedOrigin",
            Style = "position:absolute;left:1000;top:2000;width:2000;height:2000",
            CoordinateSize = "2000,2000",
            CoordinateOrigin = "1000,2000"
        };

        var outer = new DocumentFormat.OpenXml.Vml.Group(nested) {
            Id = "NativeCoverOuterOrigin",
            Style = "position:absolute;left:72;top:72;width:288;height:288",
            CoordinateSize = "4000,4000"
        };

        return new Paragraph(new Run(new Picture(outer)));
    }

    private static object BuildNativeTableOfContentsEntries(WordDocument document) {
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("BuildNativeTableOfContentsEntries", BindingFlags.NonPublic | BindingFlags.Static)!;
        return method.Invoke(null, new object[] {
            document,
            new PdfSaveOptions { IncludePageNumbers = false },
            new Dictionary<Paragraph, string>()
        })!;
    }

    private static object CreateNativeTableOfContentsEntryStyle(int relativeLevel, double contentWidth) {
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeTableOfContentsEntryStyle", BindingFlags.NonPublic | BindingFlags.Static)!;
        return method.Invoke(null, new object?[] { relativeLevel, contentWidth })!;
    }

    private static double GetPdfParagraphStyleDouble(object style, string propertyName) {
        object? value = style.GetType().GetProperty(propertyName)!.GetValue(style);
        return Assert.IsType<double>(value);
    }
}
