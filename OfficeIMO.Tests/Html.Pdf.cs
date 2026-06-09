using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;
using System.Text;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfTests {
    [Fact]
    public void Html_SaveAsPdf_SemanticProfile_ExportsThroughMarkdownPipeline() {
        var options = new HtmlPdfSaveOptions {
            Profile = HtmlPdfProfile.Semantic
        };

        byte[] pdf = """
<article>
  <h1>HTML Report</h1>
  <p><strong>OfficeIMO</strong> turns semantic HTML into PDF.</p>
  <ul><li>Markdown bridge</li><li>Shared PDF engine</li></ul>
</article>
""".SaveAsPdf(options);

        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        Assert.Contains("HTML Report", text);
        Assert.Contains("Markdown bridge", text);
        Assert.DoesNotContain(options.ConversionReport.Warnings, warning => warning.Code == "UnsupportedImage");
    }

    [Fact]
    public void Html_SaveAsPdf_SemanticProfile_ForwardsMarkdownPdfWarningsToSharedReport() {
        var markdownOptions = new MarkdownPdfSaveOptions();
        var options = new HtmlPdfSaveOptions {
            Profile = HtmlPdfProfile.Semantic,
            MarkdownHtmlOptions = HtmlToMarkdownOptions.CreateOfficeIMOProfile(),
            MarkdownPdfOptions = markdownOptions
        };

        byte[] pdf = """
<h1>Remote Asset</h1>
<p><img src="https://example.com/logo.png" alt="OfficeIMO logo"></p>
""".SaveAsPdf(options);

        Assert.True(pdf.Length > 0);
        Assert.Single(markdownOptions.Warnings, item => item.Code == "UnsupportedImage");
        PdfCore.PdfConversionWarning warning = Assert.Single(options.ConversionReport.Warnings, item => item.Code == "UnsupportedImage");
        Assert.Equal("OfficeIMO.Markdown.Pdf", warning.Converter);
        Assert.Equal("UnsupportedImage", warning.Code);

        options.ConversionReport.Clear();

        Assert.False(options.ConversionReport.HasWarnings);
        Assert.Empty(options.ConversionReport.Warnings);
    }

    [Fact]
    public void Html_SaveAsPdf_SemanticProfile_PreservesBodyCellTableAlignment() {
        var options = new HtmlPdfSaveOptions {
            Profile = HtmlPdfProfile.Semantic,
            MarkdownPdfOptions = new MarkdownPdfSaveOptions {
                PdfOptions = new PdfCore.PdfOptions {
                    PageWidth = 420,
                    PageHeight = 260,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36,
                    DefaultFontSize = 10
                }
            }
        };

        byte[] pdf = """
<table>
  <tr><th>Item</th><th>Center Qty</th><th>Right Qty</th></tr>
  <tr>
    <td>Service</td>
    <td style="text-align:center">77</td>
    <td style="text-align:right">88</td>
  </tr>
</table>
""".SaveAsPdf(options);

        PdfTextBounds centerHeader = FindPdfTextBounds(pdf, "Center", "Qty");
        PdfTextBounds rightHeader = FindPdfTextBounds(pdf, "Right", "Qty");
        PdfTextBounds centerQuantity = FindPdfTextBounds(pdf, "77");
        PdfTextBounds rightQuantity = FindPdfTextBounds(pdf, "88");

        Assert.InRange(Math.Abs(centerQuantity.Center - centerHeader.Center), 0D, 2D);
        Assert.InRange(Math.Abs(rightQuantity.Right - rightHeader.Right), 0D, 2D);
    }

    [Fact]
    public void Html_SaveAsPdf_SemanticProfile_PreservesNonUniformBodyCellTableAlignment() {
        var options = new HtmlPdfSaveOptions {
            Profile = HtmlPdfProfile.Semantic,
            MarkdownPdfOptions = new MarkdownPdfSaveOptions {
                PdfOptions = new PdfCore.PdfOptions {
                    PageWidth = 420,
                    PageHeight = 260,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36,
                    DefaultFontSize = 10
                }
            }
        };

        byte[] pdf = """
<table>
  <tr><th>Item</th><th style="text-align:center">Amount</th></tr>
  <tr><td>Subtotal</td><td style="text-align:right">125.50</td></tr>
  <tr><td>Discount</td><td style="text-align:center">10%</td></tr>
</table>
""".SaveAsPdf(options);

        PdfTextBounds amountHeader = FindPdfTextBounds(pdf, "Amount");
        PdfTextBounds subtotal = FindPdfTextBounds(pdf, "125.50");
        PdfTextBounds discount = FindPdfTextBounds(pdf, "10%");

        Assert.InRange(Math.Abs(discount.Center - amountHeader.Center), 0D, 2D);
        Assert.True(subtotal.Right > discount.Right + 20D, $"Expected right-aligned subtotal to move past the centered discount. Subtotal right: {subtotal.Right:0.##}; discount right: {discount.Right:0.##}.");
    }

    [Fact]
    public void Html_SaveAsPdf_SemanticProfile_PreservesColumnGroupTableAlignment() {
        var options = new HtmlPdfSaveOptions {
            Profile = HtmlPdfProfile.Semantic,
            MarkdownPdfOptions = new MarkdownPdfSaveOptions {
                PdfOptions = new PdfCore.PdfOptions {
                    PageWidth = 420,
                    PageHeight = 260,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36,
                    DefaultFontSize = 10
                }
            }
        };

        byte[] pdf = """
<table>
  <colgroup>
    <col>
    <col class="text-center">
    <col style="text-align:right">
  </colgroup>
  <tr><th>Item</th><th>Center Qty</th><th>Right Qty</th></tr>
  <tr><td>Service</td><td>77</td><td>88</td></tr>
</table>
""".SaveAsPdf(options);

        PdfTextBounds centerHeader = FindPdfTextBounds(pdf, "Center", "Qty");
        PdfTextBounds rightHeader = FindPdfTextBounds(pdf, "Right", "Qty");
        PdfTextBounds centerQuantity = FindPdfTextBounds(pdf, "77");
        PdfTextBounds rightQuantity = FindPdfTextBounds(pdf, "88");

        Assert.InRange(Math.Abs(centerQuantity.Center - centerHeader.Center), 0D, 2D);
        Assert.InRange(Math.Abs(rightQuantity.Right - rightHeader.Right), 0D, 2D);
    }

    [Fact]
    public void Html_SaveAsPdf_SemanticProfile_PreservesTableCellSpans() {
        var options = new HtmlPdfSaveOptions {
            Profile = HtmlPdfProfile.Semantic,
            MarkdownPdfOptions = new MarkdownPdfSaveOptions {
                PdfOptions = new PdfCore.PdfOptions {
                    CompressContentStreams = false,
                    PageWidth = 420,
                    PageHeight = 260,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36,
                    DefaultFontSize = 10
                }
            }
        };

        byte[] pdf = """
<table>
  <tr><th colspan="2">Details</th></tr>
  <tr><td rowspan="2">Service</td><td>Setup</td></tr>
  <tr><td>Support</td></tr>
</table>
""".SaveAsPdf(options);

        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        PdfTextBounds service = FindPdfTextBounds(pdf, "Service");
        PdfTextBounds setup = FindPdfTextBounds(pdf, "Setup");
        PdfTextBounds support = FindPdfTextBounds(pdf, "Support");

        Assert.Contains("Details", text);
        Assert.Contains("Support", text);
        Assert.True(setup.Left > service.Left + 40D, $"Expected Setup to render in the second column. Service left: {service.Left}; Setup left: {setup.Left}.");
        Assert.InRange(Math.Abs(support.Left - setup.Left), 0D, 2D);
    }

    [Fact]
    public void Html_SaveAsPdf_SemanticProfile_PreservesTableCellBackgroundColors() {
        var options = new HtmlPdfSaveOptions {
            Profile = HtmlPdfProfile.Semantic,
            MarkdownPdfOptions = new MarkdownPdfSaveOptions {
                PdfOptions = new PdfCore.PdfOptions {
                    CompressContentStreams = false,
                    PageWidth = 420,
                    PageHeight = 260,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36,
                    DefaultFontSize = 10
                }
            }
        };

        byte[] pdf = """
<table>
  <tr><th>Item</th><th>Total</th></tr>
  <tr><td>Service</td><td style="background-color:#663399">125.50</td></tr>
</table>
""".SaveAsPdf(options);

        string content = Encoding.ASCII.GetString(pdf);
        int fillCount = content.Split(new[] { "0.4 0.2 0.6 rg" }, StringSplitOptions.None).Length - 1;

        Assert.Equal(1, fillCount);
        Assert.Contains("125.50", PdfCore.PdfReadDocument.Load(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void Html_SaveAsPdf_SemanticProfile_PreservesTableCellTextStyles() {
        var options = new HtmlPdfSaveOptions {
            Profile = HtmlPdfProfile.Semantic,
            MarkdownPdfOptions = new MarkdownPdfSaveOptions {
                PdfOptions = new PdfCore.PdfOptions {
                    CompressContentStreams = false,
                    PageWidth = 420,
                    PageHeight = 260,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36,
                    DefaultFontSize = 10
                }
            }
        };

        byte[] pdf = """
<table>
  <tr><th>Item</th><th>Total</th></tr>
  <tr>
    <td>PlainMarker <strong>NestedBold</strong> <em>NestedItalic</em></td>
    <td style="color:#663399;font-weight:700;font-style:italic">StyledTotal</td>
  </tr>
</table>
""".SaveAsPdf(options);

        string content = Encoding.ASCII.GetString(pdf);
        int plainText = content.IndexOf("<" + Hex("PlainMarker") + ">", StringComparison.Ordinal);
        int boldText = content.IndexOf("<" + Hex("NestedBold") + ">", StringComparison.Ordinal);
        int italicText = content.IndexOf("<" + Hex("NestedItalic") + ">", StringComparison.Ordinal);
        int styledText = content.IndexOf("<" + Hex("StyledTotal") + ">", StringComparison.Ordinal);

        Assert.True(plainText >= 0, "Expected plain table text in the PDF content stream.");
        Assert.True(boldText > plainText, "Expected nested bold table text after the plain marker.");
        Assert.True(italicText > boldText, "Expected nested italic table text after the bold text.");
        Assert.True(styledText > italicText, "Expected styled total text after the nested emphasis.");

        Assert.True(content.LastIndexOf("/F2 ", boldText, StringComparison.Ordinal) > plainText, "Expected nested strong text to switch to the bold PDF font resource.");
        Assert.True(content.LastIndexOf("/F3 ", italicText, StringComparison.Ordinal) > boldText, "Expected nested emphasis text to switch to the italic PDF font resource.");
        Assert.True(content.LastIndexOf("/F4 ", styledText, StringComparison.Ordinal) > italicText, "Expected styled cell text to use the bold-italic PDF font resource.");
        Assert.True(content.LastIndexOf("0.4 0.2 0.6 rg", styledText, StringComparison.Ordinal) > italicText, "Expected styled cell text to emit the CSS text color.");
    }

    [Fact]
    public void Html_SaveAsPdf_DocumentProfile_ExportsThroughWordPipeline() {
        var options = new HtmlPdfSaveOptions {
            Profile = HtmlPdfProfile.Document,
            WordHtmlOptions = HtmlToWordOptions.CreateOfficeIMOProfile(),
            WordPdfOptions = new PdfSaveOptions()
        };

        byte[] pdf = """
<html>
  <body>
    <h1>Document HTML</h1>
    <p>Rendered through the Word HTML bridge.</p>
    <table><tr><th>Area</th><th>Status</th></tr><tr><td>HTML</td><td>PDF</td></tr></table>
  </body>
</html>
""".SaveAsPdf(options);

        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        Assert.Contains("Document HTML", text);
        Assert.Contains("Word HTML bridge", text);
    }

    [Fact]
    public void Html_SaveAsPdf_DocumentProfilePreset_PreservesPracticalHtmlFeatures() {
        string linkUri = "https://example.com/report";
        string html = CreatePracticalHtmlSample(linkUri);
        var options = HtmlPdfSaveOptions.CreateDocumentProfile();

        byte[] pdf = html.SaveAsPdf(options);

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        Assert.True(logical.PageCount >= 2);
        Assert.Contains("Practical HTML", text, StringComparison.Ordinal);
        Assert.Contains("Table marker", text, StringComparison.Ordinal);
        Assert.Contains("Second page marker", text, StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");
        Assert.Contains(logical.GetLinksByUri(linkUri), link => link.Contents == "Report link");
    }

    [Fact]
    public void HtmlPdfSaveOptions_ProfileFactories_SelectExpectedPipelines() {
        HtmlPdfSaveOptions semantic = HtmlPdfSaveOptions.CreateSemanticProfile();
        HtmlPdfSaveOptions document = HtmlPdfSaveOptions.CreateDocumentProfile();
        HtmlPdfSaveOptions trustedDocument = HtmlPdfSaveOptions.CreateTrustedDocumentProfile();

        Assert.Equal(HtmlPdfProfile.Semantic, semantic.Profile);
        Assert.NotNull(semantic.MarkdownHtmlOptions);
        Assert.NotNull(semantic.MarkdownPdfOptions);
        Assert.Equal(HtmlPdfProfile.Document, document.Profile);
        Assert.NotNull(document.WordHtmlOptions);
        Assert.NotNull(document.WordPdfOptions);
        Assert.Equal(HtmlPdfProfile.Document, trustedDocument.Profile);
        Assert.NotNull(trustedDocument.WordHtmlOptions);
        Assert.NotNull(trustedDocument.WordPdfOptions);
    }

    [Fact]
    public void HtmlPdf_ProfileContracts_CoverSupportedProfiles() {
        HtmlPdfProfileContract semantic = HtmlPdfProfileContracts.Get(HtmlPdfProfile.Semantic);
        HtmlPdfProfileContract document = HtmlPdfProfileContracts.Get(HtmlPdfProfile.Document);

        Assert.Equal(2, HtmlPdfProfileContracts.All.Count);
        Assert.Equal("html-pdf-semantic", semantic.Id);
        Assert.Contains("Markdown", semantic.Pipeline, StringComparison.Ordinal);
        Assert.Contains("semantic HTML", semantic.IntendedUse, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Not a browser renderer", semantic.UnsupportedScope, StringComparison.Ordinal);
        Assert.Equal("html-pdf-document", document.Id);
        Assert.Contains("WordDocument", document.Pipeline, StringComparison.Ordinal);
        Assert.Contains("print-oriented HTML", document.IntendedUse, StringComparison.Ordinal);
        Assert.Contains("Word HTML", document.UnsupportedScope, StringComparison.Ordinal);
        Assert.Throws<ArgumentOutOfRangeException>(() => HtmlPdfProfileContracts.Get((HtmlPdfProfile)99));
    }

    [Fact]
    public void PdfHtml_ProfileContracts_CoverSupportedProfiles() {
        PdfHtmlProfileContract semantic = PdfHtmlProfileContracts.Get(PdfHtmlProfile.Semantic);
        PdfHtmlProfileContract positioned = PdfHtmlProfileContracts.Get(PdfHtmlProfile.PositionedReview);

        Assert.Equal(2, PdfHtmlProfileContracts.All.Count);
        Assert.Equal("pdf-html-semantic", semantic.Id);
        Assert.Contains("logical model", semantic.Pipeline, StringComparison.Ordinal);
        Assert.Contains("Search", semantic.IntendedUse, StringComparison.Ordinal);
        Assert.Contains("OCR", semantic.UnsupportedScope, StringComparison.Ordinal);
        Assert.Equal("pdf-html-positioned-review", positioned.Id);
        Assert.Contains("positioned review hints", positioned.Pipeline, StringComparison.Ordinal);
        Assert.Contains("browser", positioned.IntendedUse, StringComparison.Ordinal);
        Assert.Contains("not a full PDF renderer", positioned.UnsupportedScope, StringComparison.Ordinal);
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfHtmlProfileContracts.Get((PdfHtmlProfile)99));
    }

    [Fact]
    public void Pdf_ToHtml_SemanticProfile_ExportsLogicalStructure() {
        byte[] pdf = CreateLogicalSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            LayoutOptions = new PdfCore.PdfTextLayoutOptions {
                ForceSingleColumn = true
            }
        };

        string html = PdfHtmlConverter.ToHtml(pdf, options);

        Assert.Contains("<title>Logical PDF sample</title>", html, StringComparison.Ordinal);
        Assert.Contains("<h1>Logical Heading</h1>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Logical readback marker.</p>", html, StringComparison.Ordinal);
        Assert.Contains("<ul data-pdf-list-level=\"1\"><li>Detected logical bullet</li></ul>", html, StringComparison.Ordinal);
        Assert.Contains("<table>", html, StringComparison.Ordinal);
        Assert.Contains("<th>Code</th>", html, StringComparison.Ordinal);
        Assert.Contains("<th class=\"pdf-numeric\" style=\"text-align:right\">Qty</th>", html, StringComparison.Ordinal);
        Assert.Contains("<td>A-100</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td class=\"pdf-numeric\" style=\"text-align:right\">2</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td class=\"pdf-numeric\" style=\"text-align:right\">14</td>", html, StringComparison.Ordinal);
        Assert.Equal(1, CountOccurrences(html, "A-100"));
        Assert.False(options.ConversionReport.HasWarnings);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_ExportsPageGeometryAndTextBlocks() {
        byte[] pdf = CreateLogicalSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            LayoutOptions = new PdfCore.PdfTextLayoutOptions {
                ForceSingleColumn = true
            }
        };

        string html = PdfHtmlConverter.ToHtml(PdfCore.PdfReadDocument.Load(pdf), options);

        Assert.Contains(".pdf-page{position:relative", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-page\" data-page-number=\"1\" style=\"width:420pt;height:360pt;\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-text pdf-heading\"", html, StringComparison.Ordinal);
        Assert.Contains("style=\"left:", html, StringComparison.Ordinal);
        Assert.Contains("Logical Heading", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewFragment_IncludesPositioningCss() {
        byte[] pdf = CreateLogicalSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            EmitDocumentShell = false
        };

        string html = PdfHtmlConverter.ToHtml(pdf, options);

        Assert.DoesNotContain("<!doctype html>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<style>", html, StringComparison.Ordinal);
        Assert.Contains(".pdf-page{position:relative", html, StringComparison.Ordinal);
        Assert.Contains(".pdf-text{position:absolute", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-page\" data-page-number=\"1\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_ExportsPositionedImagePlaceholders() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Canvas(canvas => canvas.Image(PdfPngTestImages.CreateRgbPng(1, 1), 40, 50, 60, 30))
            .ToBytes();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview
        };

        string html = PdfHtmlConverter.ToHtml(pdf, options);

        Assert.Contains("class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.Contains("style=\"position:absolute;left:40pt;top:50pt;width:60pt;height:30pt;\"", html, StringComparison.Ordinal);
        Assert.Contains("data-matrix=\"60 0 0 30 40 140\"", html, StringComparison.Ordinal);
        Assert.Contains("<img src=\"data:image/png;base64,", html, StringComparison.Ordinal);
        Assert.False(options.ConversionReport.HasWarnings);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_CanForceImagePlaceholders() {
        byte[] pdf = CreateImageSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            ImageExportMode = PdfHtmlImageExportMode.PlaceholderOnly
        };

        string html = PdfHtmlConverter.ToHtml(pdf, options);

        Assert.Contains("class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.Contains("<figcaption>Image:", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<img src=\"data:image/png;base64,", html, StringComparison.Ordinal);
        Assert.False(options.ConversionReport.HasWarnings);
    }

    [Fact]
    public void Pdf_ToHtml_SemanticProfile_EmbedsExtractedImageData() {
        byte[] pdf = CreateImageSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic
        };

        string html = PdfHtmlConverter.ToHtml(pdf, options);

        Assert.Contains("<figure class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.Contains("<img src=\"data:image/png;base64,", html, StringComparison.Ordinal);
        Assert.Contains("<figcaption>Image:", html, StringComparison.Ordinal);
        Assert.False(options.ConversionReport.HasWarnings);
    }

    [Fact]
    public void Pdf_ToHtml_PageRanges_ExportsSelectedPagesThroughSameBridgePackage() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .Paragraph(paragraph => paragraph.Text("First PDF page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second PDF page"))
            .ToBytes();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            PageRanges = new[] {
                PdfCore.PdfPageRange.From(2, 2)
            }
        };

        string html = PdfHtmlConverter.ToHtml(pdf, options);

        Assert.DoesNotContain("First PDF page", html, StringComparison.Ordinal);
        Assert.Contains("Second PDF page", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_PageRanges_DoesNotReapplyRangesAfterLoadingSelection() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .Paragraph(paragraph => paragraph.Text("Duplicated selected page"))
            .ToBytes();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            PageRanges = new[] {
                PdfCore.PdfPageRange.From(1, 1),
                PdfCore.PdfPageRange.From(1, 1)
            }
        };

        string html = PdfHtmlConverter.ToHtml(pdf, options);

        Assert.Equal(2, CountOrdinal(html, "<section class=\"pdf-page\""));
        Assert.Equal(2, CountOrdinal(html, "Duplicated selected page"));
    }

    [Fact]
    public void Pdf_ToHtml_PageRanges_FilterAlreadyLoadedLogicalDocument() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .Paragraph(paragraph => paragraph.Text("First logical page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second logical page"))
            .ToBytes();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf);
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            PageRanges = new[] {
                PdfCore.PdfPageRange.From(2, 2)
            }
        };

        string html = logical.ToHtml(options);

        Assert.DoesNotContain("First logical page", html, StringComparison.Ordinal);
        Assert.Contains("Second logical page", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_AccountsForRotatedPages() {
        byte[] pdf = CreateRotatedLinkAnnotationPdf(90, "https://example.com/rotated");
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        string html = PdfHtmlConverter.ToHtml(pdf, options);

        Assert.Contains("class=\"pdf-page\" data-page-number=\"1\" style=\"width:220pt;height:320pt;\"", html, StringComparison.Ordinal);
        Assert.Contains("style=\"left:38pt;top:40pt;width:22pt;height:140pt\"", html, StringComparison.Ordinal);
        Assert.Contains("href=\"https://example.com/rotated\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_FlipsCoordinatesForRotated180Pages() {
        byte[] pdf = CreateRotatedLinkAnnotationPdf(180, "https://example.com/rotated-180");
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        string html = PdfHtmlConverter.ToHtml(pdf, options);

        Assert.Contains("class=\"pdf-page\" data-page-number=\"1\" style=\"width:320pt;height:220pt;\"", html, StringComparison.Ordinal);
        Assert.Contains("style=\"left:140pt;top:38pt;width:140pt;height:22pt\"", html, StringComparison.Ordinal);
        Assert.Contains("href=\"https://example.com/rotated-180\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_LinkAnnotations_RenderUnsafeUriAsInertText() {
        const string unsafeUri = "javascript:alert(1)";
        byte[] pdf = CreateLinkAnnotationPdf(unsafeUri);
        var semanticOptions = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            IncludeLinkAnnotations = true
        };
        var positionedOptions = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        string semanticHtml = PdfHtmlConverter.ToHtml(pdf, semanticOptions);
        string positionedHtml = PdfHtmlConverter.ToHtml(pdf, positionedOptions);

        Assert.DoesNotContain("<a href=\"javascript:", semanticHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data-unsafe-href=\"javascript:alert(1)\"", semanticHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("<a href=\"javascript:", positionedHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data-unsafe-href=\"javascript:alert(1)\"", positionedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPdf_BaselineArtifacts_ExposeStableRoundTripShape() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Html.Pdf." + Guid.NewGuid().ToString("N"));
        string pdfPath = Path.Combine(directory, "practical-html.pdf");
        string htmlPath = Path.Combine(directory, "practical-html-review.html");
        string linkUri = "https://example.com/artifact";
        Directory.CreateDirectory(directory);

        try {
            CreatePracticalHtmlSample(linkUri).SaveAsPdf(pdfPath, HtmlPdfSaveOptions.CreateDocumentProfile());
            PdfHtmlConverter.SaveAsHtml(pdfPath, htmlPath, new PdfHtmlSaveOptions {
                Profile = PdfHtmlProfile.PositionedReview,
                IncludeLinkAnnotations = true
            });

            byte[] pdf = File.ReadAllBytes(pdfPath);
            string html = File.ReadAllText(htmlPath);

            Assert.True(new FileInfo(pdfPath).Length > 0);
            Assert.True(new FileInfo(htmlPath).Length > 0);
            Assert.True(PdfCore.PdfInspector.Inspect(pdf).PageCount >= 2);
            Assert.Contains("class=\"pdf-page\" data-page-number=\"1\"", html, StringComparison.Ordinal);
            Assert.Contains("class=\"pdf-link\"", html, StringComparison.Ordinal);
            Assert.Contains("href=\"" + linkUri + "\"", html, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", html, StringComparison.Ordinal);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static byte[] CreateImageSamplePdf() {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Canvas(canvas => canvas.Image(PdfPngTestImages.CreateRgbPng(1, 1), 40, 50, 60, 30))
            .ToBytes();
    }

    private static byte[] CreateLinkAnnotationPdf(string uri) {
        string escapedUri = uri.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)");
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Annots [4 0 R] >>",
            "endobj",
            "4 0 obj",
            $"<< /Type /Annot /Subtype /Link /Rect [40 160 180 182] /Contents (Unsafe link) /A << /S /URI /URI ({escapedUri}) >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static int CountOrdinal(string value, string search) {
        int count = 0;
        int index = 0;
        while (true) {
            index = value.IndexOf(search, index, StringComparison.Ordinal);
            if (index < 0) {
                return count;
            }

            count++;
            index += search.Length;
        }
    }

    private static byte[] CreateRotatedLinkAnnotationPdf(int rotationDegrees, string uri) {
        string escapedUri = uri.Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)");
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Rotate {rotationDegrees.ToString(System.Globalization.CultureInfo.InvariantCulture)} /Annots [4 0 R] >>",
            "endobj",
            "4 0 obj",
            $"<< /Type /Annot /Subtype /Link /Rect [40 160 180 182] /Contents (Rotated link) /A << /S /URI /URI ({escapedUri}) >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R >>",
            "%%EOF"
        }) + "\n";

        return System.Text.Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateLogicalSamplePdf() {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "Logical PDF sample", author: "OfficeIMO")
            .H1("Logical Heading")
            .Paragraph(paragraph => paragraph.Text("Logical readback marker."))
            .Bullets(new[] { "Detected logical bullet" })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
    }

    private static string CreatePracticalHtmlSample(string linkUri) {
        string pixel = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        return $$"""
<html>
<head>
  <style>
    table { border-collapse: collapse; }
    td, th { border: 1px solid #444; padding: 4px; }
    .page-two { break-before: page; }
  </style>
</head>
<body>
  <h1>Practical HTML</h1>
  <p><a href="{{linkUri}}">Report link</a></p>
  <p><img src="data:image/png;base64,{{pixel}}" alt="Embedded pixel" width="24" height="24"></p>
  <table>
    <tr><th>Area</th><th>Status</th></tr>
    <tr><td>Table marker</td><td>Ready</td></tr>
  </table>
  <section class="page-two"><h2>Second page marker</h2><p>Page break proof.</p></section>
</body>
</html>
""";
    }

    private static int CountOccurrences(string text, string value) {
        int count = 0;
        int index = 0;
        while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += value.Length;
        }

        return count;
    }

    private static string Hex(string text) {
        byte[] bytes = Encoding.ASCII.GetBytes(text);
        var builder = new StringBuilder(bytes.Length * 2);
        for (int i = 0; i < bytes.Length; i++) {
            builder.Append(bytes[i].ToString("X2"));
        }

        return builder.ToString();
    }

    private sealed class PdfTextBounds {
        public PdfTextBounds(double left, double right) {
            Left = left;
            Right = right;
        }

        public double Left { get; }

        public double Right { get; }

        public double Center => (Left + Right) / 2D;
    }

    private static PdfTextBounds FindPdfTextBounds(byte[] pdf, params string[] texts) {
        using PdfPigDocument document = PdfPigDocument.Open(new MemoryStream(pdf));
        var lines = document.GetPage(1)
            .GetWords()
            .GroupBy(word => Math.Round(word.BoundingBox.Bottom, 1))
            .Select(group => group.OrderBy(word => word.BoundingBox.Left).ToList())
            .ToList();

        foreach (var line in lines) {
            for (int index = 0; index <= line.Count - texts.Length; index++) {
                bool matches = true;
                for (int offset = 0; offset < texts.Length; offset++) {
                    if (!string.Equals(line[index + offset].Text, texts[offset], StringComparison.Ordinal)) {
                        matches = false;
                        break;
                    }
                }

                if (matches) {
                    double left = line.Skip(index).Take(texts.Length).Min(word => word.BoundingBox.Left);
                    double right = line.Skip(index).Take(texts.Length).Max(word => word.BoundingBox.Right);
                    return new PdfTextBounds(left, right);
                }
            }
        }

        throw new InvalidOperationException("Could not find rendered PDF text '" + string.Join(" ", texts) + "'.");
    }
}
