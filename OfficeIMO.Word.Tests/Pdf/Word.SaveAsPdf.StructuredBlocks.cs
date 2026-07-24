using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using V = DocumentFormat.OpenXml.Vml;
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
    public void SaveAsPdf_OfficeIMOEngine_Does_Not_Double_Count_CoverPage_With_Explicit_Break_For_Toc() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageExplicitBreakToc.docx");

        using WordDocument document = WordDocument.Create(docPath);
        document._document.Body!.Append(CreateNativeCoverPageBlock("Native cover title", "Native cover subtitle"));
        document.AddPageBreak();
        document.AddTableOfContent();
        document.AddPageBreak();
        document.AddParagraph("Native heading after explicit cover break").SetStyle(WordParagraphStyles.Heading1);

        object entries = BuildNativeTableOfContentsEntries(document);
        object entry = ((System.Collections.IEnumerable)entries)
            .Cast<object>()
            .Single(item => string.Equals((string)item.GetType().GetProperty("Text")!.GetValue(item)!, "Native heading after explicit cover break", StringComparison.Ordinal));
        int pageNumber = (int)entry.GetType().GetProperty("PageNumber")!.GetValue(entry)!;

        Assert.Equal(3, pageNumber);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_DoesNotAppendBlankPageAfterFinalCoverPage() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeFinalCoverPage.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeFinalCoverPage.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlock("Final native cover"));

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("Final native cover", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_DoesNotDoubleBreakAfterCoverPageWithManualPageBreak() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageManualBreak.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageManualBreak.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlock("Cover with existing separator"));
            document.AddPageBreak();
            document.AddParagraph("Body after existing separator");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Cover with existing separator", pdf.GetPage(1).Text);
        Assert.Contains("Body after existing separator", pdf.GetPage(2).Text);
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
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Ordinary_Alias_ContentControls() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeOrdinaryAliasContentControl.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeOrdinaryAliasContentControl.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.BuiltinDocumentProperties.Title = "Document property title";
            document._document.Body!.Append(new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "Title" }),
                new SdtContentBlock(
                    new Paragraph(new Run(new Text("Manual content control title"))))));
            document.AddParagraph("After ordinary content control");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Manual content control title", text);
        Assert.Contains("After ordinary content control", text);
        Assert.DoesNotContain("Document property title", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Attaches_StructuredBlock_Footnotes_Inline() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeStructuredBlockFootnote.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeStructuredBlockFootnote.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordParagraph paragraph = document.AddParagraph("Structured block footnote");
            paragraph.AddFootNote("Structured block footnote text");
            Paragraph paragraphNode = (Paragraph)paragraph._paragraph.CloneNode(true);
            paragraph._paragraph.Remove();

            document._document.Body!.Append(new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "Structured footnote block" }),
                new SdtContentBlock(paragraphNode)));
            document.AddParagraph("After structured footnote block");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        string normalized = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");
        Assert.Contains("1 Structured block footnote", normalized);
        Assert.Contains("1 Structured block footnote text", normalized);
        Assert.Contains("After structured footnote block", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Ordinary_Inline_Alias_ContentControls() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeOrdinaryInlineAliasContentControl.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeOrdinaryInlineAliasContentControl.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.BuiltinDocumentProperties.Title = "Document property title";
            document._document.Body!.Append(new Paragraph(new SdtRun(
                new SdtProperties(new SdtAlias { Val = "Title" }),
                new SdtContentRun(new Run(new Text("Manual inline content control title"))))));
            document.AddParagraph("After ordinary inline content control");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Manual inline content control title", text);
        Assert.Contains("After ordinary inline content control", text);
        Assert.DoesNotContain("Document property title", text);
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
    public void SaveAsPdf_OfficeIMOEngine_Renders_Vml_TextBox_Run_Breaks_And_Tabs() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageTextBoxBreaksTabs.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageTextBoxBreaksTabs.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlTextBoxParagraph(
                    new Paragraph(
                        new Run(
                            new Text("Line one"),
                            new Break(),
                            new Text("Line two"),
                            new TabChar(),
                            new Text("Tail"))))));
            document.AddParagraph("Native textbox break body");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Line one", text);
        Assert.Contains("Line two", text);
        Assert.Contains("Tail", text);
        Assert.DoesNotContain("Line oneLine two", text);
        Assert.Contains("Native textbox break body", text);
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
        Assert.Contains("1 1 0 rg", pageContent, StringComparison.Ordinal);
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
    public void SaveAsPdf_OfficeIMOEngine_Renders_Vml_CoverPage_Gradient_And_Fixed_Opacity() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlGradient.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlGradient.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlGradientCoverDrawingParagraph()));
            document.AddParagraph("After VML gradient cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML gradient cover", pdf.GetPage(2).Text);

        string rawPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("/ShadingType 2", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/C0 [0.267 0.447 0.769] /C1 [0.929 0.49 0.192]", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/Type /ExtGState /ca 0.5 /CA 1", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/SH1 sh", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Vml_CoverPage_Gradient_Stops() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlGradientStops.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlGradientStops.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlGradientStopsCoverDrawingParagraph()));
            document.AddParagraph("After VML gradient stops cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML gradient stops cover", pdf.GetPage(2).Text);

        string rawPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("/ShadingType 2", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/C0 [0.267 0.447 0.769] /C1 [0.929 0.49 0.192]", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/SH1 sh", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Honors_Vml_CoverPage_Fill_And_Stroke_Switches() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlSwitches.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlSwitches.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlStrokeOnlyCoverDrawingParagraph()));
            document.AddParagraph("After VML stroke-only cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML stroke-only cover", pdf.GetPage(2).Text);

        string rawPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.DoesNotContain("0.267 0.447 0.769 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.929 0.49 0.192 RG", rawPdf, StringComparison.Ordinal);
        Assert.Contains("/Type /ExtGState /ca 1 /CA 0.5", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Honors_Vml_CoverPage_Default_Stroke() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlDefaultStroke.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlDefaultStroke.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlDefaultStrokeCoverDrawingParagraph()));
            document.AddParagraph("After VML default stroke cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML default stroke cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("0 0 0 RG", pageContent, StringComparison.Ordinal);
        Assert.Contains("72 648 144 72 re", pageContent, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Vml_CoverPage_ImageData() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlImageData.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlImageData.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (var stream = new MemoryStream(CreateNativeMinimalRgbPng())) {
                imagePart.FeedData(stream);
            }

            string relationshipId = mainPart.GetIdOfPart(imagePart);
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlImageCoverDrawingParagraph(relationshipId)));
            document.AddParagraph("After VML image cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        IReadOnlyList<PdfImagePlacement> placements = PdfImageExtractor.ExtractImagePlacements(pdfPath);

        Assert.Contains(placements, placement =>
            placement.PageNumber == 1 &&
            Math.Abs(placement.Width - 144D) < 0.1D &&
            Math.Abs(placement.Height - 72D) < 0.1D);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Vml_CoverPage_Oval_And_RoundRect() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlOvalRoundRect.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlOvalRoundRect.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlOvalAndRoundRectCoverDrawingParagraph()));
            document.AddParagraph("After VML oval and roundrect cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML oval and roundrect cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.184 0.702 0.267 rg", pageContent, StringComparison.Ordinal);
        Assert.Contains("0.929 0.49 0.192 rg", pageContent, StringComparison.Ordinal);
        Assert.Contains(" c", pageContent, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Honors_Vml_CoverPage_BuiltIn_Adjustment() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlBuiltInAdjustment.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlBuiltInAdjustment.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlAdjustedBuiltInShapeParagraph()));
            document.AddParagraph("After VML adjusted built-in cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML adjusted built-in cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("132 720 m", pageContent, StringComparison.Ordinal);
        Assert.Contains("132 648 l", pageContent, StringComparison.Ordinal);
        Assert.Contains("192 684 l", pageContent, StringComparison.Ordinal);
        Assert.Contains("0.267 0.447 0.769 rg", pageContent, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Vml_CoverPage_Formula_Path() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlFormulaPath.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlFormulaPath.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlFormulaPathCoverDrawingParagraph()));
            document.AddParagraph("After VML formula cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML formula cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("132 720 m", pageContent, StringComparison.Ordinal);
        Assert.Contains("72 648 l", pageContent, StringComparison.Ordinal);
        Assert.Contains("192 684 l", pageContent, StringComparison.Ordinal);
        Assert.Contains("132 648 l", pageContent, StringComparison.Ordinal);
        Assert.Contains("72 720 l", pageContent, StringComparison.Ordinal);
        Assert.Contains("0.267 0.447 0.769 rg", pageContent, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Honors_Vml_CoverPage_Position_Alignment() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlPositionAlignment.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlPositionAlignment.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlAlignedCoverDrawingParagraph()));
            document.AddParagraph("After VML aligned cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML aligned cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("234 648 144 72 re", pageContent, StringComparison.Ordinal);
        Assert.Contains("540 0 72 36 re", pageContent, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Honors_Vml_CoverPage_Transforms() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlTransforms.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlTransforms.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlTransformedCoverDrawingParagraph()));
            document.AddParagraph("After VML transformed cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML transformed cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("0 -1 -1 0 147 745 cm", pageContent, StringComparison.Ordinal);
        Assert.Contains("-1 0 0 -1 316 720 cm", pageContent, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Honors_Vml_CoverPage_ZIndex_Order() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlZIndex.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlZIndex.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlZIndexCoverDrawingParagraph()));
            document.AddParagraph("After VML z-index cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML z-index cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        int lowerLayer = pageContent.IndexOf("0.184 0.702 0.267 rg", StringComparison.Ordinal);
        int upperLayer = pageContent.IndexOf("0.929 0.49 0.192 rg", StringComparison.Ordinal);
        Assert.True(lowerLayer >= 0, "Expected lower z-index VML fill to be emitted.");
        Assert.True(upperLayer >= 0, "Expected higher z-index VML fill to be emitted.");
        Assert.True(lowerLayer < upperLayer, "Expected VML siblings to be painted from lower z-index to higher z-index.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Honors_Wrapped_Vml_CoverPage_ZIndex_Order() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageWrappedVmlZIndex.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageWrappedVmlZIndex.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlZIndexCoverDrawingParagraph(foregroundFirst: true),
                CreateNativeVmlZIndexCoverDrawingParagraph(foregroundFirst: false)));
            document.AddParagraph("After wrapped VML z-index cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After wrapped VML z-index cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        int lowerLayer = pageContent.IndexOf("0.184 0.702 0.267 rg", StringComparison.Ordinal);
        int upperLayer = pageContent.IndexOf("0.929 0.49 0.192 rg", StringComparison.Ordinal);
        Assert.True(lowerLayer >= 0, "Expected lower z-index wrapped VML fill to be emitted.");
        Assert.True(upperLayer >= 0, "Expected higher z-index wrapped VML fill to be emitted.");
        Assert.True(lowerLayer < upperLayer, "Expected wrapped VML paragraphs to be painted from lower z-index to higher z-index.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Orders_Vml_Descendants_Across_Wrappers() {
        var low = new V.Rectangle { Id = "low", Style = "position:absolute;left:0pt;top:0pt;width:20pt;height:20pt;z-index:1" };
        var high = new V.Rectangle { Id = "high", Style = "position:absolute;left:0pt;top:0pt;width:20pt;height:20pt;z-index:10" };
        var middle = new V.Rectangle { Id = "middle", Style = "position:absolute;left:0pt;top:0pt;width:20pt;height:20pt;z-index:5" };
        var firstWrapper = new Run(new Picture(low, high));
        var secondWrapper = new Run(new Picture(middle));
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("OrderNativeVmlCoverChildren", BindingFlags.NonPublic | BindingFlags.Static)!;

        var ordered = ((IEnumerable<OpenXmlElement>)method.Invoke(null, new object[] { new OpenXmlElement[] { firstWrapper, secondWrapper } })!).ToList();

        Assert.Equal(new[] { "low", "middle", "high" }, ordered.Select(element => element.GetAttribute("id", string.Empty).Value).ToArray());
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Vml_CoverPage_Shadow() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlShadow.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlShadow.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlShadowCoverDrawingParagraph()));
            document.AddParagraph("After VML shadow cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML shadow cover", pdf.GetPage(2).Text);

        string rawPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("/Type /ExtGState /ca 0.25 /CA 0.25", rawPdf, StringComparison.Ordinal);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        int shadow = pageContent.IndexOf("78 640 120 72 re", StringComparison.Ordinal);
        int fill = pageContent.IndexOf("72 648 120 72 re", StringComparison.Ordinal);
        Assert.True(shadow >= 0, "Expected VML shadow geometry to be offset from the source shape.");
        Assert.True(fill >= 0, "Expected the source VML shape geometry to be rendered.");
        Assert.True(shadow < fill, "Expected VML shadow to be painted behind the source shape.");
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Honors_Vml_CoverPage_Stroke_Style() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlStrokeStyle.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlStrokeStyle.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlStrokeStyleCoverDrawingParagraph()));
            document.AddParagraph("After VML stroke style cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML stroke style cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.929 0.49 0.192 RG", pageContent, StringComparison.Ordinal);
        Assert.Contains("4 w", pageContent, StringComparison.Ordinal);
        Assert.Contains("1 J", pageContent, StringComparison.Ordinal);
        Assert.Contains("2 j", pageContent, StringComparison.Ordinal);
        Assert.Contains("[12 6] 0 d", pageContent, StringComparison.Ordinal);
        Assert.Contains("72 648 144 72 re S", pageContent, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Vml_CoverPage_Line_Unit_Points() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlLineUnits.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlLineUnits.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlUnitLineCoverDrawingParagraph()));
            document.AddParagraph("After VML unit line cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML unit line cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.929 0.49 0.192 RG", pageContent, StringComparison.Ordinal);
        Assert.Contains("3 w", pageContent, StringComparison.Ordinal);
        Assert.Contains("72 720 m 144 691.654 l S", pageContent, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Vml_CoverPage_Cubic_Path() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlCubicPath.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlCubicPath.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlCubicPathCoverDrawingParagraph()));
            document.AddParagraph("After VML cubic path cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML cubic path cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("72 648 m", pageContent, StringComparison.Ordinal);
        Assert.Contains("102 720 162 720 192 648 c", pageContent, StringComparison.Ordinal);
        Assert.Contains("0.267 0.447 0.769 rg", pageContent, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Vml_CoverPage_Quadratic_Path() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlQuadraticPath.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlQuadraticPath.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlQuadraticPathCoverDrawingParagraph()));
            document.AddParagraph("After VML quadratic path cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML quadratic path cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("72 648 m", pageContent, StringComparison.Ordinal);
        Assert.Contains("112 696 152 696 192 648 c", pageContent, StringComparison.Ordinal);
        Assert.Contains("0.267 0.447 0.769 rg", pageContent, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Vml_CoverPage_Relative_Cubic_Path() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlRelativeCubicPath.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeCoverPageVmlRelativeCubicPath.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            document._document.Body!.Append(CreateNativeCoverPageBlockWithChildren(
                CreateNativeVmlRelativeCubicPathCoverDrawingParagraph()));
            document.AddParagraph("After VML relative cubic path cover");

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        Assert.True(pdf.NumberOfPages >= 2);
        Assert.Contains("After VML relative cubic path cover", pdf.GetPage(2).Text);

        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("72 648 m", pageContent, StringComparison.Ordinal);
        Assert.Contains("102 720 162 720 192 648 c", pageContent, StringComparison.Ordinal);
        Assert.Contains("0.267 0.447 0.769 rg", pageContent, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Applies_Default_White_Vml_Fill() {
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("TryCreateNativeVmlShape", BindingFlags.NonPublic | BindingFlags.Static)!;
        object?[] arguments = {
            new DocumentFormat.OpenXml.Vml.Rectangle(),
            120D,
            60D,
            CreateNativeVmlFrame(120D, 60D),
            null
        };

        bool result = (bool)method.Invoke(null, arguments)!;
        OfficeShape shape = Assert.IsType<OfficeShape>(arguments[4]);

        Assert.True(result);
        Assert.Equal(OfficeColor.White, shape.FillColor);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Explicit_Vml_NoFill() {
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("TryCreateNativeVmlShape", BindingFlags.NonPublic | BindingFlags.Static)!;
        object?[] arguments = {
            new DocumentFormat.OpenXml.Vml.Rectangle {
                FillColor = "none"
            },
            120D,
            60D,
            CreateNativeVmlFrame(120D, 60D),
            null
        };

        bool result = (bool)method.Invoke(null, arguments)!;
        OfficeShape shape = Assert.IsType<OfficeShape>(arguments[4]);

        Assert.True(result);
        Assert.Null(shape.FillColor);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Explicit_Vml_NoStroke() {
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("TryCreateNativeVmlShape", BindingFlags.NonPublic | BindingFlags.Static)!;
        var stroke = new V.Stroke();
        stroke.SetAttribute(new OpenXmlAttribute("color", string.Empty, "none"));
        object?[] arguments = {
            new V.Rectangle(stroke) {
                StrokeColor = "none"
            },
            120D,
            60D,
            CreateNativeVmlFrame(120D, 60D),
            null
        };

        bool result = (bool)method.Invoke(null, arguments)!;
        OfficeShape shape = Assert.IsType<OfficeShape>(arguments[4]);

        Assert.True(result);
        Assert.Null(shape.StrokeColor);
        Assert.Equal(0D, shape.StrokeWidth);
    }

    private static object CreateNativeVmlFrame(double width, double height) {
        Type frameType = typeof(WordPdfConverterExtensions).GetNestedType("NativeVmlFrame", BindingFlags.NonPublic)!;
        return Activator.CreateInstance(
            frameType,
            BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
            binder: null,
            args: new object[] { 0D, 0D, width, height, width, height, 0D, 0D },
            culture: null)!;
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_DoesNotRenderShapeFallbackForMissingVmlImageData() {
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("ShouldRenderNativeVmlShapeFallback", BindingFlags.NonPublic | BindingFlags.Static)!;
        var imageShape = new V.Shape(
            new V.ImageData {
                RelationshipId = "rIdMissing"
            }) {
            Id = "NativeCoverMissingImage",
            Style = "position:absolute;left:0pt;top:0pt;width:120pt;height:80pt",
            FillColor = "#C1121F"
        };
        var plainShape = new V.Shape {
            Id = "NativeCoverPlainShape",
            Style = "position:absolute;left:0pt;top:0pt;width:120pt;height:80pt",
            FillColor = "#C1121F"
        };

        Assert.False((bool)method.Invoke(null, new object[] { imageShape })!);
        Assert.True((bool)method.Invoke(null, new object[] { plainShape })!);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Ignores_NonFinite_Vml_Lengths() {
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("ResolveNativeVmlLength", BindingFlags.NonPublic | BindingFlags.Static)!;

        Assert.Null(method.Invoke(null, new object?[] { "1e309pt", 1D, 1D }));
        Assert.Null(method.Invoke(null, new object?[] { "1e307in", 1D, 1D }));
        Assert.Equal(24D, Assert.IsType<double>(method.Invoke(null, new object?[] { "24pt", 1D, 1D })));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Ignores_Unsafe_Vml_TextPath_FontSize() {
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("GetNativeVmlTextPathFontSize", BindingFlags.NonPublic | BindingFlags.Static)!;

        Assert.Null(method.Invoke(null, new object[] { new V.TextPath { Style = "font-size:1e309pt" } }));
        Assert.Null(method.Invoke(null, new object[] { new V.TextPath { Style = "font-size:999999pt" } }));
        Assert.Equal(24D, Assert.IsType<double>(method.Invoke(null, new object[] { new V.TextPath { Style = "font-size:24pt" } })));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Resolves_DrawingMl_Preset_Colors() {
        var properties = new ChartShapeProperties(
            new A.SolidFill(new A.PresetColor { Val = A.PresetColorValues.Red }));
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("TryGetNativeDrawingSolidFillColor", BindingFlags.NonPublic | BindingFlags.Static)!;
        object?[] arguments = { properties, null, null };

        bool result = (bool)method.Invoke(null, arguments)!;

        Assert.True(result);
        Assert.Equal(OfficeColor.FromRgb(255, 0, 0), Assert.IsType<OfficeColor>(arguments[1]));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Applies_DrawingMl_Luminance_Transforms_In_Hsl_Space() {
        var properties = new ChartShapeProperties(
            new A.SolidFill(new A.RgbColorModelHex(
                new A.LuminanceModulation { Val = 50000 },
                new A.LuminanceOffset { Val = 20000 }) {
                Val = "4472C4"
            }));
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("TryGetNativeDrawingSolidFillColor", BindingFlags.NonPublic | BindingFlags.Static)!;
        object?[] arguments = { properties, null, null };

        bool result = (bool)method.Invoke(null, arguments)!;

        Assert.True(result);
        Assert.Equal(OfficeColor.ParseHex("#3864B2"), Assert.IsType<OfficeColor>(arguments[1]));
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
    public void SaveAsPdf_OfficeIMOEngine_Maps_Image_Watermark_To_Pdf_Watermark() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeImageWatermark.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeImageWatermark.pdf");
        string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.AddHeadersAndFooters();
            WordHeader header = RequireSectionHeader(document, 0, HeaderFooterValues.Default);
            WordWatermark watermark = header.AddWatermark(WordWatermarkStyle.Image, imagePath);
            watermark.Opacity = 0.35;
            document.AddParagraph("Native image watermark body text");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeWatermarkImageUnsupported");
        string pageContent = ReadPdfPageContent(File.ReadAllBytes(pdfPath));
        Assert.Contains("/Im", pageContent, StringComparison.Ordinal);

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Native image watermark body text", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Vml_TextBox_Run_Fonts() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeVmlTextRunFonts.docx"));
        var paragraph = new Paragraph(
            new Run(
                new RunProperties(new RunFonts { Ascii = "Courier New", HighAnsi = "Courier New" }),
                new Text("Monospace VML text")));
        Paragraph textBoxParagraph = CreateNativeVmlTextBoxParagraph(paragraph);
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("GetNativeVmlTextRuns", BindingFlags.NonPublic | BindingFlags.Static)!;

        var runs = (IReadOnlyList<TextRun>)method.Invoke(null, new object[] { document, textBoxParagraph })!;

        TextRun run = Assert.Single(runs);
        Assert.Equal("Monospace VML text", run.Text);
        Assert.Equal(PdfStandardFont.Courier, run.Font);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Honors_Vml_TextPath_And_Underline_None() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "PdfNativeVmlTextPath.docx"));
        var textPathShape = new V.Shape(
            new V.Path { AllowTextPath = true },
            new V.TextPath {
                On = true,
                FitShape = true,
                String = "TextPath cover label",
                Style = "font-family:\"Courier New\";font-size:18pt"
            }) {
            Id = "NativeCoverTextPath",
            Type = "#_x0000_t136",
            Style = "position:absolute;left:72pt;top:72pt;width:360pt;height:48pt"
        };
        var textBoxParagraph = CreateNativeVmlTextBoxParagraph(new Paragraph(
            new Run(
                new RunProperties(new Underline { Val = UnderlineValues.None }),
                new Text("Not underlined"))));
        textBoxParagraph.Append(new Run(new Picture(textPathShape)));

        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("GetNativeVmlTextRuns", BindingFlags.NonPublic | BindingFlags.Static)!;

        var runs = (IReadOnlyList<TextRun>)method.Invoke(null, new object[] { document, textBoxParagraph })!;

        TextRun plain = Assert.Single(runs, run => run.Text == "Not underlined");
        Assert.False(plain.Underline);
        TextRun textPath = Assert.Single(runs, run => run.Text == "TextPath cover label");
        Assert.Equal(PdfStandardFont.Courier, textPath.Font);
        Assert.Equal(18D, textPath.FontSize);
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
        string normalizedText = System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ");
        Assert.Contains("Native header first Native header second", normalizedText);
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

    [Theory]
    [InlineData(1L, 1L)]
    [InlineData(long.MaxValue, long.MaxValue)]
    public void SaveAsPdf_OfficeIMOEngine_NormalizesUntrustedWordChartExtents(long cx, long cy) {
        string suffix = cx == 1L ? "Tiny" : "Huge";
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChart" + suffix + "Extent.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChart" + suffix + "Extent.pdf");

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Untrusted Word chart extent", false, 360, 220);
            chart.AddPie("Passed", 3);
            chart.AddPie("Failed", 1);
            DW.Extent extent = chart.Drawing!.Inline?.Extent ?? chart.Drawing.Anchor!.Extent!;
            extent.Cx = cx;
            extent.Cy = cy;
            document.AddParagraph("After normalized chart extent");

            document.SaveAsPdf(pdfPath, new PdfSaveOptions { IncludePageNumbers = false });
        }

        Assert.Contains("After normalized chart extent", PdfTextExtractor.ExtractAllText(pdfPath));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Inline_Word_Charts_After_Text_Run() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeInlineWordChart.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeInlineWordChart.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordParagraph paragraph = document.AddParagraph("Before inline chart");
            WordChart chart = new WordChart(document, paragraph, "Inline Word PDF Chart", false, 320, 180);
            chart.AddPie("Passed", 2);
            chart.AddPie("Failed", 1);
            document.AddParagraph("After inline chart");

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Before inline chart", text);
        Assert.Contains("After inline chart", text);

        string rawPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.122 0.306 0.475 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.184 0.435 0.243 rg", rawPdf, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_OfficeIMOEngine_Ignores_Malformed_Inline_Word_Chart_References(bool relationshipPointsToImagePart) {
        string suffix = relationshipPointsToImagePart ? "WrongPart" : "MissingPart";
        string docPath = Path.Combine(_directoryWithFiles, $"PdfNativeMalformedInlineWordChart{suffix}.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, $"PdfNativeMalformedInlineWordChart{suffix}.pdf");
        const string relationshipId = "rIdMalformedInlineChart";
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordParagraph paragraph = document.AddParagraph("Before malformed inline chart");
            if (relationshipPointsToImagePart) {
                AddPngImagePart(document, relationshipId);
            }

            paragraph._paragraph!.Append(CreateMalformedInlineChartRun(relationshipId));
            document.AddParagraph("After malformed inline chart");

            document.Save();
            PdfDocumentConversionResult result = document.ToPdfDocumentResult(options);
            result.Save(pdfPath);
            Assert.Contains(result.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        }
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Before malformed inline chart", text);
        Assert.Contains("After malformed inline chart", text);
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

    private static Run CreateMalformedInlineChartRun(string relationshipId) {
        var chartReference = new ChartReference {
            Id = relationshipId
        };
        chartReference.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        chartReference.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

        return new Run(
            new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                new DW.Inline(
                    new DW.Extent {
                        Cx = 3048000L,
                        Cy = 1714500L
                    },
                    new DW.DocProperties {
                        Id = 1U,
                        Name = "malformed chart"
                    },
                    new A.Graphic(
                        new A.GraphicData(chartReference) {
                            Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                        }))));
    }

    private static void AddPngImagePart(WordDocument document, string relationshipId) {
        ImagePart imagePart = document._wordprocessingDocument.MainDocumentPart!.AddImagePart(ImagePartType.Png, relationshipId);
        byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=");
        using Stream stream = imagePart.GetStream(FileMode.Create, FileAccess.Write);
        stream.Write(png, 0, png.Length);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Word_Cartesian_DataLabels() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordCartesianDataLabels.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordCartesianDataLabels.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Bar Labels", false, 360, 220);
            chart.AddCategories(new[] { "Q1", "Q2" }.ToList());
            chart.AddBar("Actual", new[] { 10, 20 }, OfficeColor.ParseHex("#4472c4"));
            document.AddParagraph("After bar labels");

            ChartPart chartPart = (ChartPart)typeof(WordChart)
                .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
                .GetValue(chart)!;
            DataLabels labels = chartPart.ChartSpace!.Descendants<DataLabels>().First();
            labels.GetFirstChild<ShowCategoryName>()!.Val = true;
            labels.InsertBefore(new DataLabelPosition { Val = DataLabelPositionValues.Center }, labels.GetFirstChild<ShowLegendKey>());
            labels.InsertBefore(new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat {
                FormatCode = "0.0",
                SourceLinked = false
            }, labels.GetFirstChild<ShowLegendKey>());
            ValueAxis valueAxis = chartPart.ChartSpace.Descendants<ValueAxis>().First();
            valueAxis.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat>()!.FormatCode = "#,##0.0";

            object snapshot = CreateNativeWordChartSnapshot(chart);
            object layout = snapshot.GetType().GetProperty("Layout")!.GetValue(snapshot)!;
            Assert.True((bool)layout.GetType().GetProperty("ShowDataLabels")!.GetValue(layout)!);
            Assert.True((bool)layout.GetType().GetProperty("ShowDataLabelValues")!.GetValue(layout)!);
            Assert.True((bool)layout.GetType().GetProperty("ShowDataLabelCategoryNames")!.GetValue(layout)!);
            Assert.False((bool)layout.GetType().GetProperty("ShowDataLabelPercentages")!.GetValue(layout)!);
            Assert.Equal(OfficeChartDataLabelPosition.Center, layout.GetType().GetProperty("DataLabelPosition")!.GetValue(layout));
            Assert.Equal("0.0", layout.GetType().GetProperty("DataLabelNumberFormat")!.GetValue(layout));
            Assert.Equal("#,##0.0", layout.GetType().GetProperty("AxisNumberFormat")!.GetValue(layout));

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Q1; 10.0", text);
        Assert.Contains("Q2; 20.0", text);
        Assert.Contains("After bar labels", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Separates_Native_Word_Scatter_Axis_Metadata() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordScatterAxisMetadata.docx");

        using WordDocument document = WordDocument.Create(docPath);
        WordChart chart = document.AddChart("Word PDF Scatter Axis Metadata", false, 360, 220);
        chart.AddScatter("Points", new List<double> { 1, 2 }, new List<double> { 10, 20 }, OfficeColor.ParseHex("#4472c4"));

        ChartPart chartPart = (ChartPart)typeof(WordChart)
            .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
            .GetValue(chart)!;
        ScatterChart scatter = chartPart.ChartSpace!.Descendants<ScatterChart>().First();
        uint[] axisIds = scatter.Elements<AxisId>()
            .Select(axis => axis.Val!.Value)
            .ToArray();

        ValueAxis xAxis = chartPart.ChartSpace.Descendants<ValueAxis>().Single(axis => axis.AxisId!.Val!.Value == axisIds[0]);
        ValueAxis yAxis = chartPart.ChartSpace.Descendants<ValueAxis>().Single(axis => axis.AxisId!.Val!.Value == axisIds[1]);
        ApplyNumberFormat(xAxis, "0.0");
        ApplyNumberFormat(yAxis, "#,##0.00");
        xAxis.Append(CreateNativeWordChartAxisTitle("X Axis"));
        yAxis.Append(CreateNativeWordChartAxisTitle("Y Axis"));

        object snapshot = CreateNativeWordChartSnapshot(chart);
        object layout = snapshot.GetType().GetProperty("Layout")!.GetValue(snapshot)!;

        Assert.Equal("0.0", layout.GetType().GetProperty("HorizontalAxisNumberFormat")!.GetValue(layout));
        Assert.Equal("#,##0.00", layout.GetType().GetProperty("VerticalAxisNumberFormat")!.GetValue(layout));
        Assert.Equal("X Axis", layout.GetType().GetProperty("CategoryAxisTitle")!.GetValue(layout));
        Assert.Equal("Y Axis", layout.GetType().GetProperty("ValueAxisTitle")!.GetValue(layout));
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Honors_Native_Word_Deleted_Axes_And_Legend_Entries() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordHiddenAxesLegend.docx");

        using WordDocument document = WordDocument.Create(docPath);
        WordChart chart = document.AddChart("Word PDF Hidden Axes Legend", false, 360, 220);
        chart.AddCategories(new[] { "Q1", "Q2" }.ToList());
        chart.AddBar("HiddenLegend", new[] { 10, 20 }, OfficeColor.ParseHex("#4472c4"));
        chart.AddBar("VisibleLegend", new[] { 12, 24 }, OfficeColor.ParseHex("#70ad47"));
        chart.AddLegend(LegendPositionValues.Right);

        ChartPart chartPart = (ChartPart)typeof(WordChart)
            .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
            .GetValue(chart)!;
        SetNativeWordChartAxisDeleted(chartPart.ChartSpace!.Descendants<CategoryAxis>().First());
        SetNativeWordChartAxisDeleted(chartPart.ChartSpace.Descendants<ValueAxis>().First());
        Legend legend = chartPart.ChartSpace.Descendants<Legend>().First();
        legend.Append(new LegendEntry(
            new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 0U },
            new Delete { Val = true }));

        object snapshot = CreateNativeWordChartSnapshot(chart);
        object layout = snapshot.GetType().GetProperty("Layout")!.GetValue(snapshot)!;
        var data = (OfficeChartData)snapshot.GetType().GetProperty("Data")!.GetValue(snapshot)!;

        Assert.False((bool)layout.GetType().GetProperty("ShowCategoryAxis")!.GetValue(layout)!);
        Assert.False((bool)layout.GetType().GetProperty("ShowValueAxis")!.GetValue(layout)!);
        Assert.False(data.Series[0].ShowInLegend);
        Assert.True(data.Series[1].ShowInLegend);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Blank_Word_Chart_Cache_Points_As_Gaps() {
        var values = new Values(
            new NumberReference(
                new NumberingCache(
                    new PointCount { Val = 3U },
                    new NumericPoint(new NumericValue("4")) { Index = 0U },
                    new NumericPoint(new NumericValue("8")) { Index = 2U })));

        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("ExtractNativeWordChartNumberValues", BindingFlags.NonPublic | BindingFlags.Static)!;
        IReadOnlyList<double> extracted = Assert.IsAssignableFrom<IReadOnlyList<double>>(method.Invoke(null, new object?[] { values }));

        Assert.Equal(3, extracted.Count);
        Assert.Equal(4D, extracted[0]);
        Assert.True(double.IsNaN(extracted[1]));
        Assert.Equal(8D, extracted[2]);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Rejects_Partial_Word_Combo_Chart_Export() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordMixedUnsupportedChart.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordMixedUnsupportedChart.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Mixed Chart", false, 360, 220);
            chart.AddCategories(new[] { "Q1", "Q2" }.ToList());
            chart.AddBar("Actual", new[] { 10, 20 }, OfficeColor.ParseHex("#4472c4"));
            document.AddParagraph("After mixed chart");

            ChartPart chartPart = (ChartPart)typeof(WordChart)
                .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
                .GetValue(chart)!;
            PlotArea plotArea = chartPart.ChartSpace!.Descendants<PlotArea>().First();
            plotArea.InsertBefore(new BubbleChart(), plotArea.GetFirstChild<BarChart>()!);

            MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("TryCreateNativeWordChartSnapshot", BindingFlags.NonPublic | BindingFlags.Static)!;
            object?[] arguments = { chart, null, null };
            bool result = (bool)method.Invoke(null, arguments)!;

            Assert.False(result);
            Assert.Null(arguments[1]);
            Assert.Contains("not partially exported", (string?)arguments[2], StringComparison.OrdinalIgnoreCase);

            document.Save();
            PdfDocumentConversionResult conversion = document.ToPdfDocumentResult(options);
            conversion.Save(pdfPath);

            PdfConversionWarning unsupported = Assert.Single(conversion.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
            Assert.Contains("not partially exported", unsupported.Message, StringComparison.OrdinalIgnoreCase);
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After mixed chart", text);
        Assert.DoesNotContain("Word PDF Mixed Chart", text);
        Assert.DoesNotContain("Q1", text);
        Assert.DoesNotContain("Q2", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Renders_Word_Chart_AxisTitles() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartAxisTitles.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartAxisTitles.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Axis Titles", false, 360, 220);
            chart.AddCategories(new[] { "Q1", "Q2" }.ToList());
            chart.AddBar("Actual", new[] { 10, 20 }, OfficeColor.ParseHex("#4472c4"));
            chart.SetXAxisTitle("Quarter Axis");
            chart.SetYAxisTitle("Score Axis");
            document.AddParagraph("After axis titles");

            object snapshot = CreateNativeWordChartSnapshot(chart);
            object layout = snapshot.GetType().GetProperty("Layout")!.GetValue(snapshot)!;
            Assert.Equal("Quarter Axis", layout.GetType().GetProperty("CategoryAxisTitle")!.GetValue(layout));
            Assert.Equal("Score Axis", layout.GetType().GetProperty("ValueAxisTitle")!.GetValue(layout));

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("Quarter Axis", text);
        Assert.Contains("Score Axis", text);
        Assert.Contains("After axis titles", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Word_Line_Chart_NoMarkers() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordLineNoMarkers.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordLineNoMarkers.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Line No Markers", false, 360, 220);
            chart.AddCategories(new[] { "Q1", "Q2", "Q3" }.ToList());
            chart.AddLine("Trend", new List<int> { 10, 15, 20 }, OfficeColor.ParseHex("#4472c4"));
            document.AddParagraph("After line no markers");

            ChartPart chartPart = (ChartPart)typeof(WordChart)
                .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
                .GetValue(chart)!;
            LineChartSeries series = chartPart.ChartSpace!.Descendants<LineChartSeries>().First();
            series.InsertBefore(new Marker(new Symbol { Val = MarkerStyleValues.None }), series.GetFirstChild<CategoryAxisData>());

            object snapshot = CreateNativeWordChartSnapshot(chart);
            object layout = snapshot.GetType().GetProperty("Layout")!.GetValue(snapshot)!;
            Assert.False((bool)layout.GetType().GetProperty("ShowMarkers")!.GetValue(layout)!);

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After line no markers", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Word_Bar_Chart_Series_Colors() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordBarChartColors.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordBarChartColors.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Bar Colors", false, 360, 220);
            chart.AddCategories(new[] { "Q1", "Q2", "Q3" }.ToList());
            chart.AddBar("EMEA", new[] { 10, 12, 14 }, OfficeColor.Black);
            chart.AddBar("APAC", new[] { 9, 11, 15 }, OfficeColor.Black);
            chart.ApplyPalette(WordChart.WordChartPalette.ColorBlindSafe);
            document.AddParagraph("After bar colors");

            List<OfficeColor> palette = GetNativeWordChartPalette(CreateNativeWordChartSnapshot(chart));
            Assert.Equal(OfficeColor.ParseHex("#0072B2"), palette[0]);
            Assert.Equal(OfficeColor.ParseHex("#E69F00"), palette[1]);

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After bar colors", text);

        string rawPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("0 0.447 0.698 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.902 0.624 0 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Word_Chart_Area_And_Plot_Area_Colors() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartAreaColors.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartAreaColors.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Chart Area Colors", false, 360, 220);
            chart.AddCategories(new[] { "Q1", "Q2", "Q3" }.ToList());
            chart.AddBar("Actual", new[] { 10, 12, 14 }, OfficeColor.ParseHex("#4472c4"));
            document.AddParagraph("After chart area colors");

            ChartPart chartPart = (ChartPart)typeof(WordChart)
                .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
                .GetValue(chart)!;
            Chart chartElement = chartPart.ChartSpace!.GetFirstChild<Chart>()!;
            chartElement.Append(CreateNativeChartShapeProperties("FFF2CC", "7F6000"));
            chartElement.PlotArea!.Append(CreateNativeChartShapeProperties("D9EAF7", "1F4E79"));

            object snapshot = CreateNativeWordChartSnapshot(chart);
            object style = snapshot.GetType().GetProperty("Style")!.GetValue(snapshot)!;
            Assert.Equal(OfficeColor.ParseHex("#fff2cc"), style.GetType().GetProperty("BackgroundColor")!.GetValue(style));
            Assert.Equal(OfficeColor.ParseHex("#7f6000"), style.GetType().GetProperty("BorderColor")!.GetValue(style));
            Assert.Equal(OfficeColor.ParseHex("#d9eaf7"), style.GetType().GetProperty("PlotAreaBackgroundColor")!.GetValue(style));
            Assert.Equal(OfficeColor.ParseHex("#1f4e79"), style.GetType().GetProperty("PlotAreaBorderColor")!.GetValue(style));

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After chart area colors", text);

        string rawPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("1 0.949 0.8 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.498 0.376 0 RG", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.851 0.918 0.969 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.122 0.306 0.475 RG", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Word_Chart_Axis_And_Gridline_Colors() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartAxisGridColors.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartAxisGridColors.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Axis Grid Colors", false, 360, 220);
            chart.AddCategories(new[] { "Q1", "Q2", "Q3" }.ToList());
            chart.AddBar("Actual", new[] { 10, 12, 14 }, OfficeColor.ParseHex("#4472c4"));
            document.AddParagraph("After axis grid colors");

            ChartPart chartPart = (ChartPart)typeof(WordChart)
                .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
                .GetValue(chart)!;
            ValueAxis valueAxis = chartPart.ChartSpace!.Descendants<ValueAxis>().First();
            valueAxis.RemoveAllChildren<ChartShapeProperties>();
            valueAxis.Append(CreateNativeChartOutlineShapeProperties("FF0000"));
            MajorGridlines gridlines = valueAxis.GetFirstChild<MajorGridlines>() ?? new MajorGridlines();
            if (gridlines.Parent == null) {
                valueAxis.InsertAfter(gridlines, valueAxis.GetFirstChild<AxisPosition>());
            }

            gridlines.RemoveAllChildren<ChartShapeProperties>();
            gridlines.Append(CreateNativeChartOutlineShapeProperties("00FF00"));

            object snapshot = CreateNativeWordChartSnapshot(chart);
            object style = snapshot.GetType().GetProperty("Style")!.GetValue(snapshot)!;
            Assert.Equal(OfficeColor.ParseHex("#ff0000"), style.GetType().GetProperty("AxisColor")!.GetValue(style));
            Assert.Equal(OfficeColor.ParseHex("#00ff00"), style.GetType().GetProperty("GridLineColor")!.GetValue(style));

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After axis grid colors", text);

        string rawPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("1 0 0 RG", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0 1 0 RG", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Suppresses_Word_Chart_Gridlines_When_Disabled() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartNoGridlines.docx");

        using WordDocument document = WordDocument.Create(docPath);
        WordChart chart = document.AddChart("Word PDF No Gridlines", false, 360, 220);
        chart.AddCategories(new[] { "Q1", "Q2", "Q3" }.ToList());
        chart.AddBar("Actual", new[] { 10, 12, 14 }, OfficeColor.ParseHex("#4472c4"));

        ChartPart chartPart = (ChartPart)typeof(WordChart)
            .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
            .GetValue(chart)!;
        ValueAxis valueAxis = chartPart.ChartSpace!.Descendants<ValueAxis>().First();
        valueAxis.RemoveAllChildren<MajorGridlines>();

        object snapshot = CreateNativeWordChartSnapshot(chart);
        object style = snapshot.GetType().GetProperty("Style")!.GetValue(snapshot)!;

        Assert.False((bool)style.GetType().GetProperty("ShowGridLines")!.GetValue(style)!);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Word_Chart_Title_Color() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartTitleColor.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartTitleColor.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Styled Title", false, 360, 220);
            chart.AddCategories(new[] { "Q1", "Q2", "Q3" }.ToList());
            chart.AddBar("Actual", new[] { 10, 12, 14 }, OfficeColor.ParseHex("#4472c4"));
            document.AddParagraph("After chart title color");

            ChartPart chartPart = (ChartPart)typeof(WordChart)
                .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
                .GetValue(chart)!;
            A.DefaultRunProperties titleRunProperties = chartPart.ChartSpace!
                .GetFirstChild<Chart>()!
                .Title!
                .Descendants<A.DefaultRunProperties>()
                .First();
            titleRunProperties.RemoveAllChildren<A.SolidFill>();
            titleRunProperties.Append(new A.SolidFill(new A.RgbColorModelHex { Val = "CC0066" }));

            object snapshot = CreateNativeWordChartSnapshot(chart);
            object style = snapshot.GetType().GetProperty("Style")!.GetValue(snapshot)!;
            Assert.Equal(OfficeColor.ParseHex("#cc0066"), style.GetType().GetProperty("TitleColor")!.GetValue(style));

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After chart title color", text);

        string rawPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.8 0 0.4 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Word_Chart_Category_Label_Skip() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartLabelSkip.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartLabelSkip.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Label Skip", false, 360, 220);
            chart.AddCategories(Enumerable.Range(1, 12).Select(index => "Q" + index.ToString()).ToList());
            chart.AddBar("Volume", Enumerable.Range(1, 12).Select(index => index * 2).ToArray(), OfficeColor.ParseHex("#4472c4"));
            document.AddParagraph("After label skip chart");

            ChartPart chartPart = (ChartPart)typeof(WordChart)
                .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
                .GetValue(chart)!;
            CategoryAxis axis = chartPart.ChartSpace!.Descendants<CategoryAxis>().First();
            axis.Append(new TickLabelSkip { Val = 3 });

            object snapshot = CreateNativeWordChartSnapshot(chart);
            object layout = snapshot.GetType().GetProperty("Layout")!.GetValue(snapshot)!;
            Assert.Equal(4, (int)layout.GetType().GetProperty("MaximumHorizontalCategoryAxisLabels")!.GetValue(layout)!);

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After label skip chart", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Does_Not_Invent_Word_Chart_Legend() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartNoLegend.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartNoLegend.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF No Legend", false, 360, 220);
            chart.AddCategories(new[] { "Q1", "Q2", "Q3" }.ToList());
            chart.AddBar("NoLegendSeries", new[] { 10, 12, 14 }, OfficeColor.ParseHex("#4472c4"));
            document.AddParagraph("After no legend chart");

            ChartPart chartPart = (ChartPart)typeof(WordChart)
                .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
                .GetValue(chart)!;
            chartPart.ChartSpace!.Descendants<Legend>().ToList().ForEach(legend => legend.Remove());
            chartPart.ChartSpace.Descendants<DataLabels>().ToList().ForEach(labels => labels.Remove());

            object snapshot = CreateNativeWordChartSnapshot(chart);
            object layout = snapshot.GetType().GetProperty("Layout")!.GetValue(snapshot)!;
            Assert.False((bool)layout.GetType().GetProperty("ShowLegend")!.GetValue(layout)!);

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After no legend chart", text);
        Assert.DoesNotContain("NoLegendSeries", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Word_Chart_Bottom_Legend_Position() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartBottomLegend.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartBottomLegend.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Bottom Legend", false, 360, 220);
            chart.AddCategories(new[] { "Q1", "Q2", "Q3" }.ToList());
            chart.AddBar("BottomLegendSeries", new[] { 10, 12, 14 }, OfficeColor.ParseHex("#4472c4"));
            ChartPart chartPart = (ChartPart)typeof(WordChart)
                .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
                .GetValue(chart)!;
            chartPart.ChartSpace!.Descendants<Legend>().ToList().ForEach(legend => legend.Remove());
            chart.AddLegend(LegendPositionValues.Bottom);
            document.AddParagraph("After bottom legend chart");

            object snapshot = CreateNativeWordChartSnapshot(chart);
            object layout = snapshot.GetType().GetProperty("Layout")!.GetValue(snapshot)!;
            Assert.True((bool)layout.GetType().GetProperty("ShowLegend")!.GetValue(layout)!);
            Assert.Equal(OfficeChartLegendPosition.Bottom, layout.GetType().GetProperty("LegendPosition")!.GetValue(layout));

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After bottom legend chart", text);
        Assert.Contains("BottomLegendSeries", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Word_Chart_Left_Legend_Position() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartLeftLegend.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartLeftLegend.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Left Legend", false, 360, 220);
            chart.AddCategories(new[] { "Q1", "Q2", "Q3" }.ToList());
            chart.AddBar("LeftLegendSeries", new[] { 10, 12, 14 }, OfficeColor.ParseHex("#4472c4"));
            ChartPart chartPart = (ChartPart)typeof(WordChart)
                .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
                .GetValue(chart)!;
            chartPart.ChartSpace!.Descendants<Legend>().ToList().ForEach(legend => legend.Remove());
            chart.AddLegend(LegendPositionValues.Left);
            document.AddParagraph("After left legend chart");

            object snapshot = CreateNativeWordChartSnapshot(chart);
            object layout = snapshot.GetType().GetProperty("Layout")!.GetValue(snapshot)!;
            Assert.True((bool)layout.GetType().GetProperty("ShowLegend")!.GetValue(layout)!);
            Assert.Equal(OfficeChartLegendPosition.Left, layout.GetType().GetProperty("LegendPosition")!.GetValue(layout));

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After left legend chart", text);
        Assert.Contains("LeftLegendSeries", text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Word_Chart_Scheme_Series_Colors() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartSchemeColors.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartSchemeColors.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Scheme Colors", false, 360, 220);
            chart.AddCategories(new[] { "Q1", "Q2", "Q3" }.ToList());
            chart.AddBar("EMEA", new[] { 10, 12, 14 }, OfficeColor.Black);
            document.AddParagraph("After scheme chart colors");

            ChartPart chartPart = (ChartPart)typeof(WordChart)
                .GetProperty("ChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!
                .GetValue(chart)!;
            BarChartSeries series = chartPart.ChartSpace!.Descendants<BarChartSeries>().First();
            ChartShapeProperties shapeProperties = series.GetFirstChild<ChartShapeProperties>() ?? new ChartShapeProperties();
            if (shapeProperties.Parent == null) {
                series.InsertAt(shapeProperties, 2);
            }

            shapeProperties.RemoveAllChildren<A.SolidFill>();
            shapeProperties.Append(new A.SolidFill(
                new A.SchemeColor(
                    new A.LuminanceModulation() { Val = 50000 }) { Val = A.SchemeColorValues.Accent1 }));

            List<OfficeColor> palette = GetNativeWordChartPalette(CreateNativeWordChartSnapshot(chart));
            Assert.Equal(OfficeColor.ParseHex("#203864"), palette[0]);

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After scheme chart colors", text);

        string rawPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.125 0.22 0.392 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_Word_Pie_Chart_Slice_Colors() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordPieChartColors.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordPieChartColors.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Pie Colors", false, 360, 220);
            chart.AddPie("Passed", 4);
            chart.AddPie("Failed", 1);
            chart.AddPie("Skipped", 1);
            chart.ApplyPalette(WordChart.WordChartPalette.Professional, semanticOutcomes: true, applyToPies: true, applyToSeries: false);
            document.AddParagraph("After pie colors");

            List<OfficeColor> palette = GetNativeWordChartPalette(CreateNativeWordChartSnapshot(chart));
            Assert.Equal(OfficeColor.ParseHex("#2fb344"), palette[0]);
            Assert.Equal(OfficeColor.ParseHex("#f76707"), palette[1]);
            Assert.Equal(OfficeColor.ParseHex("#868e96"), palette[2]);

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After pie colors", text);

        string rawPdf = Encoding.ASCII.GetString(File.ReadAllBytes(pdfPath));
        Assert.Contains("0.184 0.702 0.267 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.969 0.404 0.027 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.525 0.557 0.588 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Uses_Word_Chart_Title_Band() {
        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartTitleBand.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeWordChartTitleBand.pdf");
        var options = new PdfSaveOptions {
            IncludePageNumbers = false
        };

        using (WordDocument document = WordDocument.Create(docPath)) {
            WordChart chart = document.AddChart("Word PDF Title Band", false, 360, 220);
            chart.AddPie("Passed", 4);
            chart.AddPie("Failed", 2);
            document.AddParagraph("After chart title band");

            object snapshot = CreateNativeWordChartSnapshot(chart);
            object layout = snapshot.GetType().GetProperty("Layout")!.GetValue(snapshot)!;
            Assert.Equal(31D, (double)layout.GetType().GetProperty("TitleTopPadding")!.GetValue(layout)!);

            document.Save();
            document.SaveAsPdf(pdfPath, options);
        }

        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.Contains("After chart title band", text);
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
                ResourcePolicy = PdfResourcePolicy.CreateTrustedHost(),
                IncludePageNumbers = false
            });
        }

        string text = PdfTextExtractor.ExtractAllText(pdfPath);
        Assert.True(
            text.Contains("✓ Native symbol text", StringComparison.Ordinal) ||
            text.Contains("✓Native symbol text", StringComparison.Ordinal),
            text);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_Preserves_System_Font_Map_In_Structured_Blocks() {
        string? fontFamily = PdfEmbeddedFontFamily.TryFromSystem("Segoe UI Symbol", out _)
            ? "Segoe UI Symbol"
            : PdfEmbeddedFontFamily.TryFromSystem("DejaVu Sans", out _)
                ? "DejaVu Sans"
                : null;
        if (fontFamily == null) {
            return;
        }

        string docPath = Path.Combine(_directoryWithFiles, "PdfNativeStructuredBlockSystemFontMap.docx");
        string pdfPath = Path.Combine(_directoryWithFiles, "PdfNativeStructuredBlockSystemFontMap.pdf");
        const string expectedText = "Structured block system font map";

        using (WordDocument document = WordDocument.Create(docPath)) {
            document.Settings.FontFamily = "Calibri";
            WordParagraph paragraph = document.AddParagraph(expectedText);
            paragraph.SetFontFamily(fontFamily);
            Paragraph paragraphNode = (Paragraph)paragraph._paragraph.CloneNode(true);
            paragraph._paragraph.Remove();

            document._document.Body!.Append(new SdtBlock(
                new SdtProperties(new SdtAlias { Val = "Structured font block" }),
                new SdtContentBlock(paragraphNode)));

            document.Save();
            document.SaveAsPdf(pdfPath, new PdfSaveOptions {
                ResourcePolicy = PdfResourcePolicy.CreateTrustedHost(),
                IncludePageNumbers = false
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(pdfPath);
        string normalizedExpectedFont = NormalizePdfFontNameForAssert(fontFamily);
        Assert.Contains(pdf.GetPage(1).Letters, letter =>
            expectedText.IndexOf(letter.Value, StringComparison.Ordinal) >= 0 &&
            letter.FontName != null &&
            NormalizePdfFontNameForAssert(letter.FontName).IndexOf(normalizedExpectedFont, StringComparison.OrdinalIgnoreCase) >= 0);
    }

    private static string NormalizePdfFontNameForAssert(string fontName) =>
        new string(fontName.Where(char.IsLetterOrDigit).ToArray());

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

    private static void SetNativeWordChartAxisDeleted(OpenXmlElement axis) {
        Delete? delete = axis.GetFirstChild<Delete>();
        if (delete == null) {
            delete = new Delete();
            axis.PrependChild(delete);
        }

        delete.Val = true;
    }

    private static List<OfficeColor> GetNativeWordChartPalette(object snapshot) {
        object style = snapshot.GetType().GetProperty("Style")!.GetValue(snapshot)!;
        return ((IEnumerable<OfficeColor>)style.GetType().GetProperty("Palette")!.GetValue(style)!).ToList();
    }

    private static ChartShapeProperties CreateNativeChartShapeProperties(string fillHex, string outlineHex) {
        return new ChartShapeProperties(
            new A.SolidFill(new A.RgbColorModelHex { Val = fillHex }),
            new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = outlineHex })));
    }

    private static ChartShapeProperties CreateNativeChartOutlineShapeProperties(string outlineHex) {
        return new ChartShapeProperties(
            new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = outlineHex })));
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
        var taggedNamedColorPanel = new DocumentFormat.OpenXml.Vml.Rectangle {
            Id = "NativeCoverTaggedNamedColorPanel",
            Style = "position:absolute;left:210;top:140;width:18;height:36",
            FillColor = "yellow [3213]",
            Stroked = false
        };

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

        return new Paragraph(new Run(new Picture(group, taggedNamedColorPanel, titleShape)));
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

    private static Paragraph CreateNativeVmlGradientCoverDrawingParagraph() {
        var fill = new DocumentFormat.OpenXml.Vml.Fill();
        fill.SetAttribute(new OpenXmlAttribute("type", string.Empty, "gradient"));
        fill.SetAttribute(new OpenXmlAttribute("color2", string.Empty, "#ed7d31"));
        fill.SetAttribute(new OpenXmlAttribute("angle", string.Empty, "90"));
        fill.SetAttribute(new OpenXmlAttribute("opacity", string.Empty, "32768F"));

        var panel = new DocumentFormat.OpenXml.Vml.Rectangle(fill) {
            Id = "NativeCoverGradientPanel",
            Style = "position:absolute;left:72;top:72;width:360;height:180",
            FillColor = "#4472c4",
            Stroked = false
        };

        return new Paragraph(new Run(new Picture(panel)));
    }

    private static Paragraph CreateNativeVmlGradientStopsCoverDrawingParagraph() {
        var fill = new DocumentFormat.OpenXml.Vml.Fill();
        fill.SetAttribute(new OpenXmlAttribute("type", string.Empty, "gradient"));
        fill.SetAttribute(new OpenXmlAttribute("color", string.Empty, "#4472c4"));
        fill.SetAttribute(new OpenXmlAttribute("colors", string.Empty, "0 #4472c4; 32768f #70ad47; 1 #ed7d31"));

        var panel = new DocumentFormat.OpenXml.Vml.Rectangle(fill) {
            Id = "NativeCoverGradientStopsPanel",
            Style = "position:absolute;left:72;top:72;width:360;height:180",
            Stroked = false
        };

        return new Paragraph(new Run(new Picture(panel)));
    }

    private static Paragraph CreateNativeVmlStrokeOnlyCoverDrawingParagraph() {
        var fill = new DocumentFormat.OpenXml.Vml.Fill();
        fill.SetAttribute(new OpenXmlAttribute("on", string.Empty, "f"));

        var stroke = new DocumentFormat.OpenXml.Vml.Stroke();
        stroke.SetAttribute(new OpenXmlAttribute("color", string.Empty, "#ed7d31"));
        stroke.SetAttribute(new OpenXmlAttribute("opacity", string.Empty, "50%"));

        var panel = new DocumentFormat.OpenXml.Vml.Rectangle(fill, stroke) {
            Id = "NativeCoverStrokeOnlyPanel",
            Style = "position:absolute;left:72;top:72;width:360;height:180",
            FillColor = "#4472c4",
            StrokeWeight = "4pt"
        };
        panel.SetAttribute(new OpenXmlAttribute("filled", string.Empty, "f"));

        return new Paragraph(new Run(new Picture(panel)));
    }

    private static Paragraph CreateNativeVmlDefaultStrokeCoverDrawingParagraph() {
        var panel = new DocumentFormat.OpenXml.Vml.Rectangle {
            Id = "NativeCoverDefaultStrokePanel",
            Style = "position:absolute;left:72;top:72;width:144;height:72",
            Filled = false
        };

        return new Paragraph(new Run(new Picture(panel)));
    }

    private static Paragraph CreateNativeVmlImageCoverDrawingParagraph(string relationshipId) {
        var imageData = new V.ImageData {
            RelationshipId = relationshipId,
            Title = "Cover image"
        };
        var shape = new V.Shape(imageData) {
            Id = "NativeCoverImageData",
            Type = "#_x0000_t75",
            Style = "position:absolute;left:72;top:72;width:144;height:72",
            Filled = false,
            Stroked = false
        };

        return new Paragraph(new Run(new Picture(shape)));
    }

    private static Paragraph CreateNativeVmlOvalAndRoundRectCoverDrawingParagraph() {
        var oval = new DocumentFormat.OpenXml.Vml.Oval {
            Id = "NativeCoverOval",
            Style = "position:absolute;left:72;top:72;width:120;height:72",
            FillColor = "#2fb344",
            Stroked = false
        };

        var roundRect = new DocumentFormat.OpenXml.Vml.RoundRectangle {
            Id = "NativeCoverRoundRect",
            Style = "position:absolute;left:216;top:72;width:144;height:72",
            FillColor = "#ed7d31",
            Stroked = false
        };
        roundRect.SetAttribute(new OpenXmlAttribute("arcsize", string.Empty, "32768F"));

        return new Paragraph(new Run(new Picture(oval, roundRect)));
    }

    private static Paragraph CreateNativeVmlAdjustedBuiltInShapeParagraph() {
        var shape = new DocumentFormat.OpenXml.Vml.Shape {
            Id = "NativeCoverAdjustedBuiltIn",
            Type = "#_x0000_t15",
            Style = "position:absolute;left:72;top:72;width:120;height:72",
            FillColor = "#4472c4",
            Stroked = false
        };
        shape.SetAttribute(new OpenXmlAttribute("adj", string.Empty, "10800"));

        return new Paragraph(new Run(new Picture(shape)));
    }

    private static Paragraph CreateNativeVmlFormulaPathCoverDrawingParagraph() {
        var shapeType = new V.Shapetype {
            Id = "_x0000_t990",
            CoordinateSize = "21600,21600",
            Adjustment = "10800",
            EdgePath = "m@0,top l,bottom right,@1 @0,bottom left,top x e"
        };
        shapeType.Append(new V.Formulas(
            new V.Formula { Equation = "val center" },
            new V.Formula { Equation = "val middle" }));

        var shape = new V.Shape {
            Id = "NativeCoverFormulaPath",
            Type = "#_x0000_t990",
            Style = "position:absolute;left:72;top:72;width:120;height:72",
            FillColor = "#4472c4",
            Stroked = false
        };
        shape.SetAttribute(new OpenXmlAttribute("adj", string.Empty, "10800"));

        return new Paragraph(new Run(new Picture(shapeType, shape)));
    }

    private static Paragraph CreateNativeVmlAlignedCoverDrawingParagraph() {
        var centered = new DocumentFormat.OpenXml.Vml.Rectangle {
            Id = "NativeCoverCentered",
            Style = "position:absolute;top:72;width:144;height:72;mso-position-horizontal:center;mso-position-horizontal-relative:page",
            FillColor = "#4472c4",
            Stroked = false
        };

        var bottomRight = new DocumentFormat.OpenXml.Vml.Rectangle {
            Id = "NativeCoverBottomRight",
            Style = "position:absolute;width:72;height:36;mso-position-horizontal:right;mso-position-horizontal-relative:page;mso-position-vertical:bottom;mso-position-vertical-relative:page",
            FillColor = "#ed7d31",
            Stroked = false
        };

        return new Paragraph(new Run(new Picture(centered, bottomRight)));
    }

    private static Paragraph CreateNativeVmlTransformedCoverDrawingParagraph() {
        var rotated = new DocumentFormat.OpenXml.Vml.Rectangle {
            Id = "NativeCoverRotated",
            Style = "position:absolute;left:72;top:72;width:100;height:50;rotation:90",
            FillColor = "#4472c4",
            Stroked = false
        };

        var flipped = new DocumentFormat.OpenXml.Vml.Rectangle {
            Id = "NativeCoverFlipped",
            Style = "position:absolute;left:216;top:72;width:100;height:50;flip:x",
            FillColor = "#ed7d31",
            Stroked = false
        };

        return new Paragraph(new Run(new Picture(rotated, flipped)));
    }

    private static Paragraph CreateNativeVmlZIndexCoverDrawingParagraph() {
        var foreground = new DocumentFormat.OpenXml.Vml.Rectangle {
            Id = "NativeCoverForeground",
            Style = "position:absolute;left:72;top:72;width:180;height:120;z-index:2",
            FillColor = "#ed7d31",
            Stroked = false
        };

        var background = new DocumentFormat.OpenXml.Vml.Rectangle {
            Id = "NativeCoverBackground",
            Style = "position:absolute;left:96;top:96;width:180;height:120;z-index:1",
            FillColor = "#2fb344",
            Stroked = false
        };

        return new Paragraph(new Run(new Picture(foreground, background)));
    }

    private static Paragraph CreateNativeVmlZIndexCoverDrawingParagraph(bool foregroundFirst) {
        var shape = foregroundFirst
            ? new DocumentFormat.OpenXml.Vml.Rectangle {
                Id = "NativeCoverForegroundWrapped",
                Style = "position:absolute;left:72;top:72;width:180;height:120;z-index:2",
                FillColor = "#ed7d31",
                Stroked = false
            }
            : new DocumentFormat.OpenXml.Vml.Rectangle {
                Id = "NativeCoverBackgroundWrapped",
                Style = "position:absolute;left:96;top:96;width:180;height:120;z-index:1",
                FillColor = "#2fb344",
                Stroked = false
            };

        return new Paragraph(new Run(new Picture(shape)));
    }

    private static Paragraph CreateNativeVmlShadowCoverDrawingParagraph() {
        var shadow = new DocumentFormat.OpenXml.Vml.Shadow();
        shadow.SetAttribute(new OpenXmlAttribute("on", string.Empty, "t"));
        shadow.SetAttribute(new OpenXmlAttribute("color", string.Empty, "#000000"));
        shadow.SetAttribute(new OpenXmlAttribute("opacity", string.Empty, "25%"));
        shadow.SetAttribute(new OpenXmlAttribute("offset", string.Empty, "6pt,8pt"));

        var panel = new DocumentFormat.OpenXml.Vml.Rectangle(shadow) {
            Id = "NativeCoverShadowPanel",
            Style = "position:absolute;left:72;top:72;width:120;height:72",
            FillColor = "#4472c4",
            Stroked = false
        };

        return new Paragraph(new Run(new Picture(panel)));
    }

    private static Paragraph CreateNativeVmlStrokeStyleCoverDrawingParagraph() {
        var fill = new DocumentFormat.OpenXml.Vml.Fill();
        fill.SetAttribute(new OpenXmlAttribute("on", string.Empty, "f"));

        var stroke = new DocumentFormat.OpenXml.Vml.Stroke();
        stroke.SetAttribute(new OpenXmlAttribute("color", string.Empty, "#ed7d31"));
        stroke.SetAttribute(new OpenXmlAttribute("dashstyle", string.Empty, "dash"));
        stroke.SetAttribute(new OpenXmlAttribute("endcap", string.Empty, "round"));
        stroke.SetAttribute(new OpenXmlAttribute("joinstyle", string.Empty, "bevel"));

        var panel = new DocumentFormat.OpenXml.Vml.Rectangle(fill, stroke) {
            Id = "NativeCoverStrokeStylePanel",
            Style = "position:absolute;left:72;top:72;width:144;height:72",
            StrokeWeight = "4pt"
        };
        panel.SetAttribute(new OpenXmlAttribute("filled", string.Empty, "f"));

        return new Paragraph(new Run(new Picture(panel)));
    }

    private static Paragraph CreateNativeVmlUnitLineCoverDrawingParagraph() {
        var line = new DocumentFormat.OpenXml.Vml.Line {
            Id = "NativeCoverUnitLine",
            Style = "position:absolute;left:72pt;top:72pt",
            From = "0pt,0pt",
            To = "1in,10mm",
            StrokeColor = "#ed7d31",
            StrokeWeight = "3pt"
        };

        return new Paragraph(new Run(new Picture(line)));
    }

    private static Paragraph CreateNativeVmlCubicPathCoverDrawingParagraph() {
        var shape = new DocumentFormat.OpenXml.Vml.Shape {
            Id = "NativeCoverCubicPath",
            Style = "position:absolute;left:72;top:72;width:120;height:72",
            FillColor = "#4472c4",
            Stroked = false
        };
        shape.SetAttribute(new OpenXmlAttribute("coordsize", string.Empty, "120,72"));
        shape.SetAttribute(new OpenXmlAttribute("path", string.Empty, "m 0,72 c 30,0 90,0 120,72 l 120,0 l 0,0 x e"));

        return new Paragraph(new Run(new Picture(shape)));
    }

    private static Paragraph CreateNativeVmlQuadraticPathCoverDrawingParagraph() {
        var shape = new DocumentFormat.OpenXml.Vml.Shape {
            Id = "NativeCoverQuadraticPath",
            Style = "position:absolute;left:72;top:72;width:120;height:72",
            FillColor = "#4472c4",
            Stroked = false
        };
        shape.SetAttribute(new OpenXmlAttribute("coordsize", string.Empty, "120,72"));
        shape.SetAttribute(new OpenXmlAttribute("path", string.Empty, "m 0,72 qb 60,0 120,72 l 120,0 l 0,0 x e"));

        return new Paragraph(new Run(new Picture(shape)));
    }

    private static Paragraph CreateNativeVmlRelativeCubicPathCoverDrawingParagraph() {
        var shape = new DocumentFormat.OpenXml.Vml.Shape {
            Id = "NativeCoverRelativeCubicPath",
            Style = "position:absolute;left:72;top:72;width:120;height:72",
            FillColor = "#4472c4",
            Stroked = false
        };
        shape.SetAttribute(new OpenXmlAttribute("coordsize", string.Empty, "120,72"));
        shape.SetAttribute(new OpenXmlAttribute("path", string.Empty, "t 0,72 v 30,-72 90,-72 120,0 r 0,-72 r -120,0 x e"));

        return new Paragraph(new Run(new Picture(shape)));
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

    private static void ApplyNumberFormat(ValueAxis axis, string formatCode) {
        DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat? numberingFormat =
            axis.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat>();
        if (numberingFormat == null) {
            numberingFormat = new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat();
            axis.Append(numberingFormat);
        }

        numberingFormat.FormatCode = formatCode;
        numberingFormat.SourceLinked = false;
    }

    private static Title CreateNativeWordChartAxisTitle(string text) =>
        new(new ChartText(new RichText(
            new A.BodyProperties(),
            new A.ListStyle(),
            new A.Paragraph(new A.Run(new A.Text(text))))));

    private static double GetPdfParagraphStyleDouble(object style, string propertyName) {
        object? value = style.GetType().GetProperty(propertyName)!.GetValue(style);
        return Assert.IsType<double>(value);
    }

    private static byte[] CreateNativeMinimalRgbPng() => PdfPngTestImages.CreateRgbPng(1, 1);

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_NativeVmlImageReader_Rejects_Images_Above_Limit() {
        using var stream = new NativeVmlGeneratedStream(WordPdfConverterExtensions.NativeVmlImageMaxBytes + 1L);

        bool result = WordPdfConverterExtensions.TryReadNativeVmlImageBytes(stream, WordPdfConverterExtensions.NativeVmlImageMaxBytes, out byte[]? imageBytes);

        Assert.False(result);
        Assert.Null(imageBytes);
        Assert.True(stream.BytesRead <= WordPdfConverterExtensions.NativeVmlImageMaxBytes + 81920L);
    }

    [Fact]
    public void SaveAsPdf_OfficeIMOEngine_NativeVmlImageReader_Reads_Images_Within_Limit() {
        byte[] source = { 1, 2, 3, 4 };
        using var stream = new MemoryStream(source);

        bool result = WordPdfConverterExtensions.TryReadNativeVmlImageBytes(stream, WordPdfConverterExtensions.NativeVmlImageMaxBytes, out byte[]? imageBytes);

        Assert.True(result);
        Assert.Equal(source, imageBytes);
    }

    private sealed class NativeVmlGeneratedStream : Stream {
        private long _remaining;

        public NativeVmlGeneratedStream(long length) {
            _remaining = length;
        }

        public long BytesRead { get; private set; }

        public override bool CanRead => true;

        public override bool CanSeek => false;

        public override bool CanWrite => false;

        public override long Length => throw new NotSupportedException();

        public override long Position {
            get => BytesRead;
            set => throw new NotSupportedException();
        }

        public override void Flush() {
        }

        public override int Read(byte[] buffer, int offset, int count) {
            if (_remaining == 0L) {
                return 0;
            }

            int bytesToRead = (int)Math.Min(count, _remaining);
            for (int i = offset; i < offset + bytesToRead; i++) {
                buffer[i] = 1;
            }

            _remaining -= bytesToRead;
            BytesRead += bytesToRead;
            return bytesToRead;
        }

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

}
