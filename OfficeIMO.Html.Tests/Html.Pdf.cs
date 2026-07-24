using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Tests.Pdf;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Pdf;
using System.Text;
using System.Threading.Tasks;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfTests {
    [Fact]
    public void HtmlToPdf_SkipsEmptySvgWithAlternativeText() {
        string svg = Convert.ToBase64String(Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' width='10' height='10'></svg>"));
        string html = "<img src='data:image/svg+xml;base64," + svg +
            "' alt='Empty vector'><p>After empty vector</p>";

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdf();

        Assert.Contains(
            "After empty vector",
            PdfCore.PdfReadDocument.Open(pdf).ExtractText(),
            StringComparison.Ordinal);
    }

    [Fact]
    public void Html_DirectOutputs_UseOneSharedOptionsShape() {
        const string html = "<main><h1>Quarterly report</h1><p>Direct HTML rendering.</p></main>";
        var options = new HtmlPdfSaveOptions {
            ViewportWidth = 640D,
            Margins = HtmlRenderMargins.All(24D),
            Scale = 1D
        };

        byte[] png = HtmlConversionDocument.Parse(html).ToPng(options);
        string svg = HtmlConversionDocument.Parse(html).ToSvg(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(options);

        Assert.True(png.Length > 8);
        Assert.True(pdf.Length > 8);
        Assert.StartsWith("<svg", svg, StringComparison.Ordinal);
        Assert.Equal(HtmlRenderMode.Paged, options.Mode);
    }

    [Fact]
    public void Html_ToPdfResult_ReturnsDiagnosticsWithoutMutatingReusableOptions() {
        const string html = "<p><img src='https://example.invalid/missing.png'>Report</p>";
        var options = new HtmlPdfSaveOptions();

        PdfCore.PdfDocumentConversionResult first = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(options);
        PdfCore.PdfDocumentConversionResult second = OfficeIMO.Html.HtmlConversionDocument.Parse("<p>Clean</p>").ToPdfDocumentResult(options);

        Assert.Contains(first.Report.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
        Assert.DoesNotContain(second.Report.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ExternalImagePending);
        Assert.Equal(HtmlRenderMode.Paged, options.Mode);
    }

    [Fact]
    public async Task Html_Pdf_BytesDocumentFileAndStream_AreConsistent() {
        const string html = "<article><h1>API contract</h1><p>One direct renderer.</p></article>";
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pdf");
        try {
            byte[] bytes = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf();
            PdfCore.PdfDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocument();
            using var stream = new MemoryStream();
            await OfficeIMO.Html.HtmlConversionDocument.Parse(html).SaveAsPdfAsync(stream);
            OfficeIMO.Html.HtmlConversionDocument.Parse(html).SaveAsPdf(path);

            Assert.Equal((byte)'%', bytes[0]);
            Assert.True(document.ToBytes().Length > 8);
            Assert.Equal((byte)'%', stream.ToArray()[0]);
            Assert.True(new FileInfo(path).Length > 8L);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void Html_OfficeProjections_AreExplicitTargets() {
        const string html = "<article><h1>Projection</h1><p>Explicit conversion.</p></article>";

        using OfficeIMO.Word.WordDocument word = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument();
        OfficeIMO.Markdown.MarkdownDoc markdown = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToMarkdownDocument();

        Assert.NotNull(word);
        Assert.Contains("Projection", markdown.ToMarkdown(), StringComparison.Ordinal);
    }
    [Fact]
    public void PdfHtml_ProfileContracts_CoverSupportedProfiles() {
        PdfHtmlProfileContract semantic = PdfHtmlProfileContracts.Get(PdfHtmlProfile.Semantic);
        PdfHtmlProfileContract positioned = PdfHtmlProfileContracts.Get(PdfHtmlProfile.PositionedReview);

        Assert.Equal(2, PdfHtmlProfileContracts.All.Count);
        Assert.Equal(HtmlConversionProfile.Semantic, semantic.SharedProfile);
        Assert.Equal("pdf-html-semantic", semantic.Id);
        Assert.Contains("logical model", semantic.Pipeline, StringComparison.Ordinal);
        Assert.Contains("Search", semantic.IntendedUse, StringComparison.Ordinal);
        Assert.Contains("OCR", semantic.UnsupportedScope, StringComparison.Ordinal);
        Assert.Contains("tables", semantic.PreservedSignals);
        Assert.Contains("export-summary", semantic.OutputArtifacts);
        Assert.Contains("no-editable-office-reconstruction", semantic.RendererBoundaries);
        Assert.Equal(HtmlConversionProfile.PositionedReview, positioned.SharedProfile);
        Assert.Equal("pdf-html-positioned-review", positioned.Id);
        Assert.Contains("positioned review hints", positioned.Pipeline, StringComparison.Ordinal);
        Assert.Contains("browser", positioned.IntendedUse, StringComparison.Ordinal);
        Assert.Contains("not a full PDF renderer", positioned.UnsupportedScope, StringComparison.Ordinal);
        Assert.Contains("image-placements", positioned.ReviewSignals);
        Assert.Contains("unsafe-link-sanitization", positioned.DiagnosticGuarantees);
        Assert.Contains("no-full-graphics-renderer", positioned.RendererBoundaries);
        Assert.Throws<ArgumentOutOfRangeException>(() => PdfHtmlProfileContracts.Get((PdfHtmlProfile)99));
    }

    [Fact]
    public void Pdf_ToHtml_SemanticProfile_ExportsLogicalStructure() {
        byte[] pdf = CreateLogicalSamplePdf();
        var layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic
        };

        string html = PdfCore.PdfLogicalDocument.Load(pdf, layoutOptions).ToHtml(options);

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
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_ExportsPageGeometryAndTextBlocks() {
        byte[] pdf = CreateLogicalSamplePdf();
        var layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview
        };

        string html = PdfCore.PdfLogicalDocument.Load(pdf, layoutOptions).ToHtml(options);

        Assert.Contains(".pdf-page{position:relative", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\" style=\"width:420pt;height:360pt;\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-text pdf-heading\"", html, StringComparison.Ordinal);
        Assert.Contains("style=\"left:", html, StringComparison.Ordinal);
        Assert.Contains("Logical Heading", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtmlResult_PositionedReviewProfile_ReportsExportSummary() {
        byte[] pdf = CreatePdfHtmlSummarySamplePdf("https://example.com/summary");
        var layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        PdfHtmlConversionResult result = PdfCore.PdfLogicalDocument.Load(pdf, layoutOptions).ToHtmlResult(options);

        Assert.False(result.Report.HasWarnings);
        Assert.Contains("Logical Heading", result.Value, StringComparison.Ordinal);
        Assert.Equal(PdfHtmlProfile.PositionedReview, result.Summary.Profile);
        Assert.Equal("pdf-html-positioned-review", result.Summary.ProfileId);
        Assert.Equal(1, result.Summary.SourcePageCount);
        Assert.Equal(1, result.Summary.RenderedPageCount);
        Assert.Equal(new[] { 1 }, result.Summary.PageNumbers);
        Assert.True(result.Summary.TextBlockCount > 0);
        Assert.True(result.Summary.HeadingCount > 0);
        Assert.True(result.Summary.ListItemCount > 0);
        Assert.Equal(0, result.Summary.TableCount);
        Assert.True(result.Summary.ImageCount > 0);
        Assert.True(result.Summary.ImagePlacementCount > 0);
        Assert.True(result.Summary.LinkCount > 0);
        Assert.Equal(0, result.Summary.WarningCount);
        Assert.True(result.Summary.EmitsDocumentShell);
        Assert.Equal(PdfHtmlImageExportMode.EmbeddedDataUri, result.Summary.ImageExportMode);
        Assert.Contains("positioned", result.Summary.FidelityContract, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("not a full PDF renderer", result.Summary.UnsupportedScope, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtmlResult_PositionedReviewProfile_RendersOutlinesAsNavigationMetadata() {
        byte[] pdf = CreateOutlineSamplePdf();
        var layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview
        };

        PdfHtmlConversionResult result = PdfCore.PdfLogicalDocument.Load(pdf, layoutOptions).ToHtmlResult(options);

        Assert.Contains("class=\"pdf-outline\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("aria-label=\"PDF outline\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-outline-count=\"3\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-rendered-outline-count=\"3\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-outline-level=\"1\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-outline-level=\"2\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("href=\"#pdf-page-1\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("href=\"#pdf-page-2\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("id=\"pdf-page-1\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("id=\"pdf-page-2\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("Executive summary", result.Value, StringComparison.Ordinal);
        Assert.Contains("Risk posture", result.Value, StringComparison.Ordinal);
        Assert.Contains("Appendix", result.Value, StringComparison.Ordinal);
        Assert.Equal(3, result.Summary.OutlineCount);
        Assert.Equal(3, result.Summary.RenderedOutlineCount);
    }

    [Fact]
    public void Pdf_ToHtmlResult_CanSuppressOutlineNavigation() {
        byte[] pdf = CreateOutlineSamplePdf();
        var layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeOutlines = false
        };

        PdfHtmlConversionResult result = PdfCore.PdfLogicalDocument.Load(pdf, layoutOptions).ToHtmlResult(options);

        Assert.DoesNotContain("class=\"pdf-outline\"", result.Value, StringComparison.Ordinal);
        Assert.Equal(3, result.Summary.OutlineCount);
        Assert.Equal(0, result.Summary.RenderedOutlineCount);
    }

    [Fact]
    public void Pdf_ToHtmlResult_ReportsAcroFormXfaAsInertReviewMetadata() {
        byte[] pdf = CreateAcroFormXfaPdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview
        };

        PdfHtmlConversionResult result = PdfCore.PdfLogicalDocument.Load(pdf).ToHtmlResult(options);

        Assert.Contains("class=\"pdf-xfa-notice\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-xfa-packet-count=\"2\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-xfa-packet-names=\"template,datasets\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("does not render or fill XFA", result.Value, StringComparison.Ordinal);
        Assert.True(result.Summary.HasAcroFormXfa);
        Assert.Equal(2, result.Summary.AcroFormXfaPacketCount);
        Assert.Equal(2, result.Summary.AcroFormXfaStreamCount);
        Assert.True(result.Summary.AcroFormXfaPayloadByteCount > 0);
        Assert.Equal(1, result.Summary.WarningCount);
        PdfCore.PdfConversionWarning warning = Assert.Single(result.Report.Warnings, item => item.Code == "AcroFormXfaDetected");
        Assert.Equal("OfficeIMO.Html.Pdf", warning.Converter);
        Assert.Contains("does not render or fill XFA", warning.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtmlResult_SnapshotsConversionReportWhenOptionsAreReused() {
        byte[] imagePdf = CreateImageSamplePdf();
        byte[] textPdf = CreateLogicalSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            MaxEmbeddedImageBytes = 0
        };

        PdfHtmlConversionResult imageResult = PdfCore.PdfLogicalDocument.Load(imagePdf).ToHtmlResult(options);
        PdfCore.PdfConversionWarning warning = Assert.Single(imageResult.Report.Warnings, item => item.Code == "ImageDataTooLarge");
        Assert.Equal("OfficeIMO.Html.Pdf", warning.Converter);

        PdfHtmlConversionResult textResult = PdfCore.PdfLogicalDocument.Load(textPdf).ToHtmlResult(options);

        Assert.Single(imageResult.Report.Warnings, item => item.Code == "ImageDataTooLarge");
        Assert.False(textResult.Report.HasWarnings);
    }

    [Fact]
    public void Pdf_ToHtmlResult_PageRanges_PreserveSourcePageCountAndSelectedFormFields() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .TextField("FirstPageField", width: 120, value: "first")
            .PageBreak()
            .TextField("SecondPageField", width: 120, value: "second")
            .ToBytes();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            PageRanges = new[] {
                PdfCore.PdfPageRange.From(2, 2)
            }
        };

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf);
        PdfHtmlConversionResult result = logical.ToHtmlResult(options);

        Assert.Equal(2, result.Summary.SourcePageCount);
        Assert.Equal(1, result.Summary.RenderedPageCount);
        Assert.Equal(new[] { 2 }, result.Summary.PageNumbers);
        Assert.Equal(1, result.Summary.FormFieldCount);
        Assert.Equal(1, result.Summary.FormWidgetCount);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewFragment_IncludesPositioningCss() {
        byte[] pdf = CreateLogicalSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            EmitDocumentShell = false
        };

        string html = PdfCore.PdfLogicalDocument.Load(pdf).ToHtml(options);

        Assert.DoesNotContain("<!doctype html>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<style>", html, StringComparison.Ordinal);
        Assert.Contains(".pdf-page{position:relative", html, StringComparison.Ordinal);
        Assert.Contains(".pdf-text{position:absolute", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\"", html, StringComparison.Ordinal);
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

        string html = PdfCore.PdfLogicalDocument.Load(pdf).ToHtml(options);

        Assert.Contains("class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.Contains("style=\"position:absolute;left:40pt;top:50pt;width:60pt;height:30pt;\"", html, StringComparison.Ordinal);
        Assert.Contains("data-matrix=\"60 0 0 30 40 140\"", html, StringComparison.Ordinal);
        Assert.Contains("<img src=\"data:image/png;base64,", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_PositionedReviewProfile_CanForceImagePlaceholders() {
        byte[] pdf = CreateImageSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            ImageExportMode = PdfHtmlImageExportMode.PlaceholderOnly
        };

        string html = PdfCore.PdfLogicalDocument.Load(pdf).ToHtml(options);

        Assert.Contains("class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.Contains("<figcaption>Image:", html, StringComparison.Ordinal);
        Assert.DoesNotContain("<img src=\"data:image/png;base64,", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_LogicalDocumentPageRanges_UsesUniqueAnchorsForDuplicatePageSelections() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                CreateOutlineFromHeadings = true,
                PageWidth = 320,
                PageHeight = 220,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36
            })
            .H1("Repeated Page")
            .ToBytes();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf);
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            PageRanges = new[] {
                PdfCore.PdfPageRange.From(1, 1),
                PdfCore.PdfPageRange.From(1, 1)
            }
        };

        string html = logical.ToHtml(options);

        Assert.Contains("id=\"pdf-page-1-1\"", html, StringComparison.Ordinal);
        Assert.Contains("id=\"pdf-page-1-2\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("id=\"pdf-page-1\"", html, StringComparison.Ordinal);
        Assert.Contains("href=\"#pdf-page-1-1\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtml_SemanticProfile_EmbedsExtractedImageData() {
        byte[] pdf = CreateImageSamplePdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic
        };

        string html = PdfCore.PdfLogicalDocument.Load(pdf).ToHtml(options);

        Assert.Contains("<figure class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.Contains("<img src=\"data:image/png;base64,", html, StringComparison.Ordinal);
        Assert.Contains("<figcaption>Image:", html, StringComparison.Ordinal);
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

        string html = PdfCore.PdfLogicalDocument.Load(pdf).ToHtml(options);

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

        string html = PdfCore.PdfLogicalDocument.Load(pdf).ToHtml(options);

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

        string html = PdfCore.PdfLogicalDocument.Load(pdf).ToHtml(options);

        Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\" style=\"width:220pt;height:320pt;\"", html, StringComparison.Ordinal);
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

        string html = PdfCore.PdfLogicalDocument.Load(pdf).ToHtml(options);

        Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\" style=\"width:320pt;height:220pt;\"", html, StringComparison.Ordinal);
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

        string semanticHtml = PdfCore.PdfLogicalDocument.Load(pdf).ToHtml(semanticOptions);
        string positionedHtml = PdfCore.PdfLogicalDocument.Load(pdf).ToHtml(positionedOptions);

        Assert.DoesNotContain("<a href=\"javascript:", semanticHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data-unsafe-href=\"javascript:alert(1)\"", semanticHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("<a href=\"javascript:", positionedHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data-unsafe-href=\"javascript:alert(1)\"", positionedHtml, StringComparison.Ordinal);
    }

    [Fact]
    public void Pdf_ToHtmlResult_ReportsActiveActionDiagnosticsWithoutPayloads() {
        byte[] pdf = CreateActiveContentDiagnosticsPdf();
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        PdfHtmlConversionResult result = PdfCore.PdfLogicalDocument.Load(pdf).ToHtmlResult(options);

        Assert.True(result.Summary.HasOpenAction);
        Assert.True(result.Summary.HasCatalogActions);
        Assert.True(result.Summary.HasPageActions);
        Assert.True(result.Summary.HasAnnotationActions);
        Assert.True(result.Summary.HasActiveContent);
        Assert.Equal(5, result.Summary.PotentiallyUnsafeActionCount);
        Assert.Equal(2, result.Summary.JavaScriptActionCount);
        Assert.Equal(1, result.Summary.LaunchActionCount);
        Assert.Equal(1, result.Summary.SubmitFormActionCount);
        Assert.Equal(1, result.Summary.CatalogActionCount);
        Assert.Equal(1, result.Summary.PageActionCount);
        Assert.Equal(1, result.Summary.SelectedPageActionCount);
        Assert.Equal(3, result.Summary.AnnotationActionCount);
        Assert.Equal(3, result.Summary.SelectedAnnotationActionCount);
        Assert.DoesNotContain("app.alert", result.Value, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("tool.exe", result.Value, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("https://example.com/submit", result.Value, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlPdf_BaselineArtifacts_ExposeStableRoundTripShape() {
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Html.Pdf." + Guid.NewGuid().ToString("N"));
        string pdfPath = Path.Combine(directory, "practical-html.pdf");
        string htmlPath = Path.Combine(directory, "practical-html-review.html");
        string linkUri = "https://example.com/artifact";
        Directory.CreateDirectory(directory);

        try {
            OfficeIMO.Html.HtmlConversionDocument.Parse(CreatePracticalHtmlSample(linkUri)).SaveAsPdf(pdfPath, new HtmlPdfSaveOptions());
            PdfCore.PdfLogicalDocument.Load(pdfPath).SaveAsHtml(htmlPath, new PdfHtmlSaveOptions {
                Profile = PdfHtmlProfile.PositionedReview,
                IncludeLinkAnnotations = true
            });

            byte[] pdf = File.ReadAllBytes(pdfPath);
            string html = File.ReadAllText(htmlPath);

            Assert.True(new FileInfo(pdfPath).Length > 0);
            Assert.True(new FileInfo(htmlPath).Length > 0);
            Assert.True(PdfCore.PdfInspector.Inspect(pdf).PageCount >= 2);
            Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\"", html, StringComparison.Ordinal);
            Assert.Contains("class=\"pdf-link\"", html, StringComparison.Ordinal);
            Assert.Contains("href=\"" + linkUri + "\"", html, StringComparison.Ordinal);
            Assert.Contains("data:image/png;base64,", html, StringComparison.Ordinal);
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static byte[] CreateAcroFormXfaPdf() {
        const string template = "<template/>";
        const string datasets = "<datasets/>";
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm 5 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 240 180] /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Fields [] /XFA [(template) 6 0 R (datasets) 7 0 R] >>",
            "endobj",
            "6 0 obj",
            "<< /Length " + template.Length + " >>",
            "stream",
            template,
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /Length " + datasets.Length + " >>",
            "stream",
            datasets,
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 8 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
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

    private static byte[] CreateActiveContentDiagnosticsPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /Fit] /Names << /JavaScript << /Names [(Open) 6 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Contents 4 0 R /Annots [5 0 R 9 0 R] /AA << /O 7 0 R >> >>",
            "endobj",
            "4 0 obj",
            "<< /Length 0 >>",
            "stream",
            "",
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Link /Rect [40 160 180 182] /Contents (Action link) /A << /S /Launch /F (tool.exe) >> /AA << /E 8 0 R >> >>",
            "endobj",
            "6 0 obj",
            "<< /S /JavaScript /JS (app.alert('catalog')) >>",
            "endobj",
            "7 0 obj",
            "<< /S /JavaScript /JS (app.alert('page')) >>",
            "endobj",
            "8 0 obj",
            "<< /S /SubmitForm /F (https://example.com/submit) >>",
            "endobj",
            "9 0 obj",
            "<< /Type /Annot /Subtype /Screen /Rect [40 110 180 150] /A << /S /RichMedia >> >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 10 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
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

    private static byte[] CreateOutlineSamplePdf() {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                CreateOutlineFromHeadings = true,
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .H1("Executive summary")
            .Paragraph(paragraph => paragraph.Text("Summary body."))
            .H2("Risk posture")
            .Paragraph(paragraph => paragraph.Text("Risk body."))
            .PageBreak()
            .H1("Appendix")
            .Paragraph(paragraph => paragraph.Text("Appendix body."))
            .ToBytes();
    }

    private static byte[] CreatePdfHtmlSummarySamplePdf(string linkUri) {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 420,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "PDF to HTML summary sample", author: "OfficeIMO")
            .H1("Logical Heading")
            .Paragraph(paragraph => paragraph.Text("Logical readback marker."))
            .Bullets(new[] { "Detected logical bullet" })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" }
            }, style: new PdfCore.PdfTableStyle {
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .Image(PdfPngTestImages.CreateRgbPng(1, 1), 24, 24, PdfCore.PdfAlign.Left, null, null, 6, 0, null, linkUri, "Summary link")
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
