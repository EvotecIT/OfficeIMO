using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.AsciiDoc;
using OfficeIMO.AsciiDoc.Pdf;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Latex;
using OfficeIMO.Latex.Pdf;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word.Pdf;
using DrawingCore = OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;
using TransformGroup = DocumentFormat.OpenXml.Drawing.TransformGroup;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConversionScenarioManifestTests {
    [Fact]
    public void PdfConversionManifest_CoversEverySupportedConversionPathWithObservableProof() {
        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(GetManifestPath()));
        JsonElement root = document.RootElement;
        Assert.Equal(2, root.GetProperty("version").GetInt32());

        JsonElement scenarios = root.GetProperty("scenarios");
        Assert.NotEmpty(scenarios.EnumerateArray());

        var seenIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonElement scenario in scenarios.EnumerateArray()) {
            string id = RequireString(scenario, "id");
            Assert.True(seenIds.Add(id), "Scenario ids must be unique. Duplicate: " + id);

            string path = RequireString(scenario, "path");
            Assert.False(string.IsNullOrWhiteSpace(path));
            Assert.Equal("supported", RequireString(scenario, "status"));
            Assert.False(string.IsNullOrWhiteSpace(RequireString(scenario, "converter")));
            Assert.False(string.IsNullOrWhiteSpace(RequireString(scenario, "sourceFormat")));
            Assert.False(string.IsNullOrWhiteSpace(RequireString(scenario, "targetFormat")));
            Assert.NotEmpty(ReadStringArray(scenario, "sourceFeatures"));
            Assert.True(ReadStringArray(scenario, "visualReviewFiles").Count > 0, "Scenario " + id + " needs at least one review artifact.");

            JsonElement proof = scenario.GetProperty("proof");
            Assert.True(proof.GetProperty("hash").GetBoolean(), "Scenario " + id + " must include hash proof.");
            Assert.NotEmpty(ReadStringArray(proof, "textMarkers"));
            Assert.NotEmpty(ReadStringArray(proof, "logicalSignals"));
            Assert.True(proof.GetProperty("visualPages").GetInt32() > 0, "Scenario " + id + " must declare visual page evidence.");
            Assert.NotEmpty(ReadStringArray(proof, "validatorEvidence"));
        }

        AssertPremiumQualityContract(root.GetProperty("qualityContract"), seenIds);
        AssertConverterCatalog(root.GetProperty("converterCatalog"), seenIds);
        AssertCompositionRoutes(root.GetProperty("compositionRoutes"));
    }

    [Fact]
    public void PdfEditableOfficeManifestUsesCanonicalSemanticWordArtifact() {
        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(GetManifestPath()));
        JsonElement scenario = document.RootElement
            .GetProperty("scenarios")
            .EnumerateArray()
            .Single(element => element.GetProperty("path").GetString() == "pdf-to-editable-office");
        IReadOnlyList<string> reviewFiles = ReadStringArray(scenario, "visualReviewFiles");

        Assert.Contains("pdf-semantic-import-word.docx", reviewFiles);
        Assert.DoesNotContain("pdf-table-import-word.docx", reviewFiles);
        Assert.Contains("pdf-table-import-excel.xlsx", reviewFiles);
        Assert.Contains("pdf-table-import-powerpoint.pptx", reviewFiles);
    }

    [Fact]
    public void AsciiDocDirectAdapter_ProducesManifestedLossAwareProof() {
        const string source =
            "= AsciiDoc direct PDF proof\n" +
            ":author: OfficeIMO\n\n" +
            "== Reusable route\n" +
            "Semantic route marker with stem:[x^2].\n\n" +
            "* Native AsciiDoc parser\n" +
            "* Shared Markdown PDF renderer\n";

        PdfCore.PdfDocumentConversionResult result = AsciiDocDocument.Parse(source).Document.ToPdfDocumentResult();
        byte[] pdf = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("AsciiDoc direct PDF proof", text, StringComparison.Ordinal);
        Assert.Contains("Semantic route marker", text, StringComparison.Ordinal);
        Assert.Contains(result.Warnings, warning => warning.Converter == "OfficeIMO.AsciiDoc.Pdf" && warning.Code == "ADOCMD103");
        Assert.Contains(result.Warnings, warning => warning.Details.TryGetValue("stage", out string? stage) && stage == "semantic-projection");

        WriteReviewArtifact("asciidoc-direct-semantic-document.pdf", pdf);
    }

    [Fact]
    public void LatexDirectAdapter_ProducesManifestedLossAwareProof() {
        const string source =
            "\\documentclass{article}\n" +
            "\\title{LaTeX direct PDF proof}\n" +
            "\\author{OfficeIMO}\n" +
            "\\begin{document}\n" +
            "\\maketitle\n" +
            "\\section{Reusable route}\n" +
            "Semantic route marker with citation \\cite{officeimo} and math $x^2$.\n" +
            "\\end{document}\n";

        PdfCore.PdfDocumentConversionResult result = LatexDocument.Parse(source).Document.ToPdfDocumentResult();
        byte[] pdf = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("LaTeX direct PDF proof", text, StringComparison.Ordinal);
        Assert.Contains("Semantic route marker", text, StringComparison.Ordinal);
        Assert.Contains(result.Warnings, warning => warning.Converter == "OfficeIMO.Latex.Pdf" && warning.Code == "LATEXMD101");
        Assert.Contains(result.Warnings, warning => warning.Code == "LATEXMD102");

        WriteReviewArtifact("latex-direct-semantic-document.pdf", pdf);
    }

    [Fact]
    public void HtmlDirectRenderer_ProducesManifestedReviewProof() {
        const string linkUri = "https://example.com/pdf-conversion-manifest";
        byte[] pdf = HtmlConversionDocument.Parse(CreatePracticalHtmlSample(linkUri)).ToPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        Assert.True(pdf.Length > 0);
        Assert.True(logical.PageCount >= 2);
        Assert.Contains(logical.TextBlocks, block => block.Text.IndexOf("Practical HTML", StringComparison.Ordinal) >= 0);
        Assert.Contains(logical.TextBlocks, block => block.Text.IndexOf("Second page marker", StringComparison.Ordinal) >= 0);
        Assert.Contains(logical.GetLinksByUri(linkUri), link => link.Contents == "Report link");
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");

        WriteReviewArtifact("practical-html.pdf", pdf);
    }

    [Fact]
    public void PdfLogicalAndHtmlProfiles_ProduceManifestedReadbackProof() {
        byte[] pdf = CreateLogicalProofPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        PdfHtmlConversionResult result = PdfHtmlConverterExtensions.ToHtmlResult(logical, options);
        string html = result.Value;
        byte[] activeContentPdf = CreatePdfToHtmlActiveContentProofPdf();
        PdfHtmlConversionResult activeContentResult = PdfHtmlConverterExtensions.ToHtmlResult(PdfCore.PdfLogicalDocument.Load(activeContentPdf), new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        });
        byte[] xfaPdf = CreateReaderXfaFormCorpusPdf();
        PdfHtmlConversionResult xfaResult = PdfHtmlConverterExtensions.ToHtmlResult(PdfCore.PdfLogicalDocument.Load(xfaPdf), new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview
        });

        Assert.Equal(1, logical.PageCount);
        Assert.Contains(logical.Headings, heading => heading.Text == "Logical Heading");
        Assert.Contains(logical.ListItems, item => item.Text == "Detected logical bullet");
        Assert.NotEmpty(logical.Tables);
        Assert.Contains(logical.GetLinksByUri("https://example.com/logical-proof"), link => link.Contents == "Logical PDF sample");
        Assert.Contains(logical.Images, image => image.Width > 0D && image.Height > 0D);
        Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-text pdf-heading\"", html, StringComparison.Ordinal);
        Assert.Contains("Logical Heading", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-outline\"", html, StringComparison.Ordinal);
        Assert.Contains("href=\"#pdf-page-1\"", html, StringComparison.Ordinal);
        Assert.Equal(1, result.Summary.RenderedPageCount);
        Assert.Equal(3, result.Summary.OutlineCount);
        Assert.Equal(3, result.Summary.RenderedOutlineCount);
        Assert.True(result.Summary.HasAnnotationActions);
        Assert.Equal(1, result.Summary.RenderedSafeUriLinkCount);
        Assert.Equal(0, result.Summary.PotentiallyUnsafeActionCount);
        Assert.True(activeContentResult.Summary.HasActiveContent);
        Assert.Equal(4, activeContentResult.Summary.PotentiallyUnsafeActionCount);
        Assert.Equal(2, activeContentResult.Summary.JavaScriptActionCount);
        Assert.Equal(1, activeContentResult.Summary.LaunchActionCount);
        Assert.Equal(1, activeContentResult.Summary.SubmitFormActionCount);
        Assert.DoesNotContain("app.alert", activeContentResult.Value, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("tool.exe", activeContentResult.Value, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("https://example.com/submit", activeContentResult.Value, StringComparison.OrdinalIgnoreCase);
        Assert.True(xfaResult.Summary.HasAcroFormXfa);
        Assert.Equal(2, xfaResult.Summary.AcroFormXfaPacketCount);
        Assert.Equal(2, xfaResult.Summary.AcroFormXfaStreamCount);
        Assert.True(xfaResult.Summary.AcroFormXfaPayloadByteCount > 0);
        Assert.Contains("class=\"pdf-xfa-notice\"", xfaResult.Value, StringComparison.Ordinal);
        Assert.Contains("data-xfa-packet-names=\"template,datasets\"", xfaResult.Value, StringComparison.Ordinal);
        Assert.Single(xfaResult.Report.Warnings, warning => warning.Code == "AcroFormXfaDetected");

        var summary = new {
            scenario = "pdf-to-html-positioned-review",
            positioned = result.Summary,
            activeContent = activeContentResult.Summary,
            activeContentPayloadPolicy = "Catalog, page, and annotation actions are counted as inert diagnostics; action payloads are not emitted into review HTML.",
            xfaForm = xfaResult.Summary,
            xfaWarningCodes = xfaResult.Report.Warnings.Select(warning => warning.Code).Distinct(StringComparer.Ordinal).ToArray()
        };

        WriteReviewArtifact("pdf-to-html-logical-source.pdf", pdf);
        WriteReviewArtifact("pdf-to-html-positioned-review.html", Encoding.UTF8.GetBytes(html));
        WriteReviewArtifact("pdf-to-html-active-content-source.pdf", activeContentPdf);
        WriteReviewArtifact("pdf-to-html-active-content-positioned-review.html", Encoding.UTF8.GetBytes(activeContentResult.Value));
        WriteReviewArtifact("pdf-to-html-xfa-form-source.pdf", xfaPdf);
        WriteReviewArtifact("pdf-to-html-xfa-form-positioned-review.html", Encoding.UTF8.GetBytes(xfaResult.Value));
        WriteReviewArtifact("pdf-to-html-positioned-review-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions { WriteIndented = true }));
    }

    [Fact]
    public void TypographyProfile_ProducesMultilingualBusinessReportProof() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] pdf;
        try {
            pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                    CompressContentStreams = false,
                    CompressEmbeddedFonts = false,
                    PageWidth = 520,
                    PageHeight = 420,
                    MarginLeft = 36,
                    MarginRight = 36,
                    MarginTop = 36,
                    MarginBottom = 36
                })
                .UseFontFamily("OfficeIMO Multilingual", fontPath)
                .Header(header => header.Text("Q2 multilingual revenue report"))
                .H1("Q2 multilingual revenue report")
                .Paragraph(paragraph => paragraph.Text("Executive summary: Zażółć gęślą jaźń Łódź."))
                .Paragraph(paragraph => paragraph.Text("Regional notes: Ελλάδα Athens pipeline; Київ renewal forecast."))
                .Table(new[] {
                    new[] { "Region", "Signal", "Status" },
                    new[] { "Polska", "Łódź", "Ready" },
                    new[] { "Ελλάδα", "Athens", "Ready" },
                    new[] { "Україна", "Київ", "Watch" }
                }, style: new PdfCore.PdfTableStyle {
                    HeaderRowCount = 1,
                    CellPaddingX = 6,
                    CellPaddingY = 4
                })
                .Footer(footer => footer.Text("Generated proof {page}/{pages}"))
                .ToBytes();
        } catch (ArgumentException exception) when (exception.Message.Contains("not covered by the embedded TrueType font", StringComparison.Ordinal)) {
            return;
        }

        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains("Q2 multilingual revenue report", text, StringComparison.Ordinal);
        Assert.Contains("Zażółć gęślą jaźń", text, StringComparison.Ordinal);
        Assert.Contains("Ελλάδα", text, StringComparison.Ordinal);
        Assert.Contains("Київ", text, StringComparison.Ordinal);

        WriteReviewArtifact("multilingual-business-report.pdf", pdf);
    }

    [Fact]
    public void MarkdownInvoiceStatement_ProducesManifestedReviewProof() {
        byte[] pdf = OfficeIMO.Markdown.MarkdownReader.Parse(CreateInvoiceStatementMarkdown()).ToPdf(new MarkdownPdfSaveOptions {
            ApplyDefaultTheme = true,
            Title = "OfficeIMO invoice statement proof",
            Subject = "Invoice and statement conversion proof"
        });

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        Assert.True(logical.PageCount >= 1);
        Assert.Contains("Invoice Statement INV-2026-0042", text, StringComparison.Ordinal);
        Assert.Contains("Managed PDF conversion review", text, StringComparison.Ordinal);
        Assert.Contains("Subtotal", text, StringComparison.Ordinal);
        Assert.Contains("Amount due", text, StringComparison.Ordinal);
        Assert.Contains(logical.Tables, table => table.Rows.Count >= 4);
        Assert.Contains(logical.ListItems, item => item.Text.Contains("Payment terms", StringComparison.Ordinal));

        WriteReviewArtifact("markdown-invoice-statement.pdf", pdf);
    }

    [Fact]
    public void RtfPdfRoundtrip_ProducesManifestedReviewProof() {
        RtfDocument source = CreateRtfRoundtripDocument();
        byte[] pdf = source.ToPdf();
        PdfCore.PdfTextLayoutOptions layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };
        PdfCore.PdfReadDocument read = PdfCore.PdfReadDocument.Open(pdf);
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, layoutOptions);
        RtfDocument imported = logical.ToRtfDocument(new PdfRtfReadOptions());
        string importedRtf = imported.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        string importedText = string.Join("\n", imported.Paragraphs.Select(paragraph => paragraph.ToPlainText()));

        Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(pdf, 0, 5), StringComparison.Ordinal);
        Assert.Equal("RTF PDF Roundtrip Gate", imported.Info.Title);
        Assert.Contains("RTF PDF Roundtrip Gate", read.ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Rich run marker", read.ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Table cell marker", read.ExtractText(), StringComparison.Ordinal);
        Assert.Contains("Imported second page marker", read.ExtractText(), StringComparison.Ordinal);
        Assert.True(read.Pages.Count >= 2);
        Assert.Contains(logical.TextBlocks, block => block.Text.Contains("RTF PDF Roundtrip Gate", StringComparison.Ordinal));
        Assert.Contains(logical.TextBlocks, block => block.Text.Contains("Table cell marker", StringComparison.Ordinal));
        Assert.Contains("RTF PDF Roundtrip Gate", importedText, StringComparison.Ordinal);
        Assert.Contains("Rich run marker", importedText, StringComparison.Ordinal);
        Assert.Contains("Table cell marker", importedText, StringComparison.Ordinal);
        Assert.Contains("Imported second page marker", importedText, StringComparison.Ordinal);
        Assert.Contains(@"\rtf1", importedRtf, StringComparison.Ordinal);

        var summary = new {
            scenario = "rtf-pdf-roundtrip",
            pdfBytes = pdf.Length,
            read.Pages.Count,
            logical.PageCount,
            headingCount = logical.Headings.Count,
            tableCount = logical.Tables.Count,
            importedParagraphCount = imported.Paragraphs.Count,
            importedTitle = imported.Info.Title,
            acceptedDegradation = "PDF to RTF import is semantic logical text reconstruction, not lossless fixed-layout visual reconstruction."
        };

        WriteReviewArtifact("rtf-roundtrip-source.pdf", pdf);
        WriteReviewArtifact("rtf-roundtrip-imported.rtf", Encoding.UTF8.GetBytes(importedRtf));
        WriteReviewArtifact("rtf-roundtrip-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions { WriteIndented = true }));
    }

    [Fact]
    public void PdfRewritePreservationMatrix_ProducesManifestedReviewProof() {
        byte[] source = PdfRewritePreservationTestSupport.BuildPreservationProofPdf();
        byte[] updated = PdfCore.PdfMetadataEditor.UpdateMetadata(source, title: "Updated preservation title");

        PdfCore.PdfRewritePreservationOptions options = new PdfCore.PdfRewritePreservationOptions()
            .AllowMetadataChanges("Title")
            .RequireTextMarkers("PreservationMarker", "SecondPageMarker");
        PdfCore.PdfRewritePreservationReport preserved = PdfCore.PdfRewritePreservation.AssertPreserved(source, updated, options);
        PdfCore.PdfRewritePreservationReport detectedLoss = PdfCore.PdfRewritePreservation.Assess(
            source,
            PdfCore.PdfPageEditor.DeletePages(source, 2),
            new PdfCore.PdfRewritePreservationOptions().RequireTextMarkers("SecondPageMarker"));

        Assert.True(preserved.IsPreserved);
        Assert.False(detectedLoss.IsPreserved);
        Assert.Contains(detectedLoss.Issues, issue => issue.Feature == "PageCount");
        Assert.Contains(detectedLoss.Issues, issue => issue.Feature == "TextMarker");

        PdfCore.PdfRewritePreservationMatrixReport matrix = PdfCore.PdfRewritePreservationMatrix.AssertExpected(
            PdfRewritePreservationMatrixScenarioSupport.BuildPremiumFeatureScenarios());

        Assert.True(matrix.Passed, matrix.Summary);
        Assert.Contains(matrix.Entries, entry =>
            entry.Id == "source-structure-drift-detected" &&
            entry.ActualClassification == PdfCore.PdfRewritePreservationMatrixClassification.PreservationFailed &&
            entry.PreservationReport is not null &&
            entry.PreservationReport.Issues.Any(issue => issue.Feature == "SourceStructure.XrefStreams"));
        Assert.Contains(matrix.Entries, entry =>
            entry.Id == "optional-content-drift-detected" &&
            entry.ActualClassification == PdfCore.PdfRewritePreservationMatrixClassification.PreservationFailed &&
            entry.PreservationReport is not null &&
            entry.PreservationReport.Issues.Any(issue => issue.Feature == "OptionalContent.Groups[0].Name"));
        Assert.Contains(matrix.Entries, entry =>
            entry.Id == "form-fill-safe" &&
            entry.ActualClassification == PdfCore.PdfRewritePreservationMatrixClassification.RewriteSafe);
        Assert.Contains(matrix.Entries, entry =>
            entry.Id == "forms-full-rewrite-blocked" &&
            entry.ActualClassification == PdfCore.PdfRewritePreservationMatrixClassification.Blocked &&
            entry.FailureMessage is not null &&
            entry.FailureMessage.Contains("PDF form fields are not supported for rewriting", StringComparison.Ordinal));
        Assert.Contains(matrix.Entries, entry =>
            entry.Id == "tagged-full-rewrite-blocked" &&
            entry.ActualClassification == PdfCore.PdfRewritePreservationMatrixClassification.Blocked &&
            entry.FailureMessage is not null &&
            entry.FailureMessage.Contains("PDF tagged content structure is not supported for rewriting", StringComparison.Ordinal));
        Assert.Contains(matrix.Entries, entry =>
            entry.Id == "active-content-full-rewrite-blocked" &&
            entry.ActualClassification == PdfCore.PdfRewritePreservationMatrixClassification.Blocked &&
            entry.FailureMessage is not null &&
            entry.FailureMessage.Contains("PDF active content is not supported for rewriting", StringComparison.Ordinal));
        Assert.Contains(matrix.Entries, entry =>
            entry.Id == "signed-full-rewrite-blocked" &&
            entry.ActualClassification == PdfCore.PdfRewritePreservationMatrixClassification.Blocked &&
            entry.FailureMessage is not null &&
            entry.FailureMessage.Contains("Signed PDF files are not supported for rewriting", StringComparison.Ordinal));

        PdfCore.PdfRewritePreservationMatrixSummary matrixSummary = matrix.ToSummary();
        var summary = new {
            scenario = "pdf-rewrite-preservation-matrix",
            preserved = preserved.IsPreserved,
            matrix = new {
                passed = matrixSummary.Passed,
                matrixSummary.Summary,
                rows = matrixSummary.Rows
            },
            source = new {
                preserved.Original.PageCount,
                preserved.Original.LinkAnnotationCount,
                namedDestinations = preserved.Original.NamedDestinations.Count,
                attachments = preserved.Original.Attachments.Count,
                outputIntents = preserved.Original.OutputIntents.Count,
                preserved.Original.HasXmpMetadata,
                preserved.Original.CatalogPageMode,
                preserved.Original.CatalogPageLayout,
                preserved.Original.CatalogLanguage
            },
            updated = new {
                preserved.Rewritten.PageCount,
                preserved.Rewritten.LinkAnnotationCount,
                namedDestinations = preserved.Rewritten.NamedDestinations.Count,
                attachments = preserved.Rewritten.Attachments.Count,
                outputIntents = preserved.Rewritten.OutputIntents.Count,
                preserved.Rewritten.HasXmpMetadata,
                preserved.Rewritten.CatalogPageMode,
                preserved.Rewritten.CatalogPageLayout,
                preserved.Rewritten.CatalogLanguage,
                preserved.Rewritten.Metadata.Title
            },
            detectedIssues = detectedLoss.Issues.Select(issue => new {
                issue.Feature,
                issue.Expected,
                issue.Actual,
                issue.Message
            }).ToArray()
        };

        byte[] summaryBytes = JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions { WriteIndented = true });
        using (JsonDocument summaryDocument = JsonDocument.Parse(summaryBytes)) {
            JsonElement root = summaryDocument.RootElement;
            Assert.True(root.GetProperty("matrix").GetProperty("passed").GetBoolean());
            JsonElement rows = root.GetProperty("matrix").GetProperty("rows");
            Assert.Equal(PdfRewritePreservationMatrixScenarioSupport.PremiumScenarioCount, rows.GetArrayLength());
            Assert.Contains(rows.EnumerateArray(), row =>
                RequireString(row, "Id") == "metadata-update-safe" &&
                RequireString(row, "ActualClassification") == "RewriteSafe");
            Assert.Contains(rows.EnumerateArray(), row =>
                RequireString(row, "Id") == "source-structure-drift-detected" &&
                RequireString(row, "ActualClassification") == "PreservationFailed" &&
                HasIssueFeature(row, "SourceStructure.XrefStreams"));
            Assert.Contains(rows.EnumerateArray(), row =>
                RequireString(row, "Id") == "optional-content-drift-detected" &&
                RequireString(row, "ActualClassification") == "PreservationFailed" &&
                HasIssueFeature(row, "OptionalContent.Groups[0].Name"));
            Assert.Contains(rows.EnumerateArray(), row =>
                RequireString(row, "Id") == "form-fill-safe" &&
                RequireString(row, "ActualClassification") == "RewriteSafe");
            Assert.Contains(rows.EnumerateArray(), row =>
                RequireString(row, "Id") == "forms-full-rewrite-blocked" &&
                RequireString(row, "ActualClassification") == "Blocked" &&
                RequireString(row, "FailureMessage").Contains("PDF form fields are not supported for rewriting", StringComparison.Ordinal));
            Assert.Contains(rows.EnumerateArray(), row =>
                RequireString(row, "Id") == "tagged-full-rewrite-blocked" &&
                RequireString(row, "ActualClassification") == "Blocked" &&
                RequireString(row, "FailureMessage").Contains("PDF tagged content structure is not supported for rewriting", StringComparison.Ordinal));
            Assert.Contains(rows.EnumerateArray(), row =>
                RequireString(row, "Id") == "active-content-full-rewrite-blocked" &&
                RequireString(row, "ActualClassification") == "Blocked" &&
                RequireString(row, "FailureMessage").Contains("PDF active content is not supported for rewriting", StringComparison.Ordinal));
            Assert.Contains(rows.EnumerateArray(), row =>
                RequireString(row, "Id") == "signed-full-rewrite-blocked" &&
                RequireString(row, "ActualClassification") == "Blocked" &&
                RequireString(row, "FailureMessage").Contains("Signed PDF files are not supported for rewriting", StringComparison.Ordinal));
        }

        WriteReviewArtifact("pdf-rewrite-preservation-source.pdf", source);
        WriteReviewArtifact("pdf-rewrite-preservation-updated.pdf", updated);
        WriteReviewArtifact("pdf-rewrite-preservation-summary.json", summaryBytes);
    }

    [Fact]
    public void PdfRedactionRemovalProof_ProducesManifestedReviewProof() {
        PdfRedactionProofResult proof = PdfRedactionProofTestSupport.BuildAndVerifyRedactionRemovalProof();
        PdfCore.PdfRedactionVerificationReport unredactedCheck = PdfCore.PdfRedactionVerification.Verify(
            proof.Source,
            PdfRedactionProofTestSupport.CreateVerificationOptions());

        Assert.True(proof.Plan.HasMatches);
        Assert.True(proof.Verification.IsVerified);
        Assert.False(unredactedCheck.IsVerified);
        Assert.Contains(proof.Plan.Matches, match => match.Text != null && match.Text.Contains("PAY-SECRET-2026", StringComparison.Ordinal));
        Assert.DoesNotContain("PAY-SECRET-2026", proof.Verification.ExtractedText, StringComparison.Ordinal);
        Assert.Contains("Visible compliance marker", proof.Verification.ExtractedText, StringComparison.Ordinal);
        Assert.Contains("Public summary marker", proof.Verification.ExtractedText, StringComparison.Ordinal);

        var summary = new {
            scenario = "pdf-redaction-removal-proof",
            plan = new {
                proof.Plan.HasMatches,
                matchCount = proof.Plan.Matches.Count,
                matchedKinds = proof.Plan.Matches.Select(match => match.Kind.ToString()).Distinct(StringComparer.Ordinal).ToArray()
            },
            area = new {
                proof.Area.PageNumber,
                proof.Area.X,
                proof.Area.Y,
                proof.Area.Width,
                proof.Area.Height,
                proof.Area.Label
            },
            verification = new {
                proof.Verification.IsVerified,
                proof.Verification.RawPdfBytesChecked,
                retainedMarkers = new[] { "Visible compliance marker", "Public summary marker" },
                removedMarkers = new[] { "Sensitive payroll token", "PAY-SECRET-2026" }
            },
            unredactedIssues = unredactedCheck.Issues.Select(issue => new {
                issue.Feature,
                issue.Marker,
                issue.Message
            }).ToArray()
        };

        WriteReviewArtifact("pdf-redaction-removal-source.pdf", proof.Source);
        WriteReviewArtifact("pdf-redaction-removal-redacted.pdf", proof.Redacted);
        WriteReviewArtifact("pdf-redaction-removal-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions { WriteIndented = true }));
    }

    [Fact]
    public void PdfFormAppearanceSemantics_ProducesManifestedReviewProof() {
        PdfFormAppearanceProofResult proof = PdfFormAppearanceProofTestSupport.BuildFormAppearanceProof();

        PdfCore.PdfFormField name = proof.FilledInfo.FormFieldsByName["Name"];
        PdfCore.PdfFormField country = proof.FilledInfo.FormFieldsByName["Country"];
        PdfCore.PdfFormField acceptTerms = proof.FilledInfo.FormFieldsByName["AcceptTerms"];
        PdfCore.PdfFormField paymentMethod = proof.FilledInfo.FormFieldsByName["Payment.Method"];
        PdfCore.PdfFormField notes = proof.FilledInfo.FormFieldsByName["Notes"];
        PdfCore.PdfFormField code = proof.FilledInfo.FormFieldsByName["Code"];
        PdfCore.PdfFormField countries = proof.FilledInfo.FormFieldsByName["Countries"];
        string countryDisplay = Assert.Single(country.SelectedOptions).DisplayText;

        Assert.Equal(false, proof.FilledInfo.AcroFormNeedAppearances);
        Assert.Equal("Visible Value", name.Value);
        Assert.Equal("PL", country.Value);
        Assert.Equal("Poland", countryDisplay);
        Assert.Equal("Yes", acceptTerms.Value);
        Assert.True(acceptTerms.IsCheckBox);
        Assert.Equal("Yes", Assert.Single(acceptTerms.Widgets).AppearanceState);
        Assert.Contains("Yes", Assert.Single(acceptTerms.Widgets).NormalAppearanceStates);
        Assert.Equal("Wire", paymentMethod.Value);
        Assert.True(paymentMethod.IsRadioButton);
        Assert.Contains(paymentMethod.Widgets, widget => widget.AppearanceState == "Wire");
        Assert.Contains(paymentMethod.Widgets, widget => widget.AppearanceState == "Off");
        Assert.True(notes.IsMultiline);
        Assert.Equal(PdfCore.PdfFormFieldTextAlignment.Center, notes.TextAlignment);
        Assert.Equal("Line one\nLine two", notes.Value);
        Assert.True(code.IsComb);
        Assert.Equal(4, code.MaxLength);
        Assert.Equal("ZX91", code.Value);
        Assert.True(countries.AllowsMultipleSelection);
        Assert.Equal(new[] { "DE", "US" }, countries.Values);
        Assert.Equal(new[] { "Germany", "United States" }, countries.SelectedOptions.Select(option => option.DisplayText).ToArray());
        Assert.Contains("<56697369626C652056616C7565> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<506F6C616E64> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<4C696E65206F6E65> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<4C696E652074776F> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<5A> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<58> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<39> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<31> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.DoesNotContain("<5A583931> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<4765726D616E79> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("<556E6974656420537461746573> Tj", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("/AS /Yes", proof.FilledRaw, StringComparison.Ordinal);
        Assert.Contains("/AS /Wire", proof.FilledRaw, StringComparison.Ordinal);
        Assert.False(proof.FlattenedInfo.HasForms);
        Assert.Contains("<56697369626C652056616C7565> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<506F6C616E64> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<4C696E65206F6E65> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<4C696E652074776F> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<5A> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<58> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<39> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<31> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.DoesNotContain("<5A583931> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<4765726D616E79> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("<556E6974656420537461746573> Tj", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("1.25 w", proof.FlattenedAppearanceText, StringComparison.Ordinal);
        Assert.Contains("Wire selected", proof.FlattenedAppearanceText, StringComparison.Ordinal);

        var summary = new {
            scenario = "pdf-form-appearance-semantics",
            filled = new {
                proof.FilledInfo.HasReadableFormFields,
                proof.FilledInfo.AcroFormNeedAppearances,
                fields = proof.FilledInfo.FormFields.Select(field => new {
                    field.Name,
                    field.Kind,
                    field.Value,
                    field.Values,
                    field.Flags,
                    field.MaxLength,
                    field.IsMultiline,
                    field.IsComb,
                    field.AllowsMultipleSelection,
                    field.TextAlignment,
                    selectedDisplayText = field.SelectedOptions.Select(option => option.DisplayText).ToArray(),
                    widgetAppearanceStates = field.Widgets.Select(widget => widget.AppearanceState).ToArray(),
                    widgetNormalAppearanceStates = field.Widgets.Select(widget => widget.NormalAppearanceStates.ToArray()).ToArray()
                }).ToArray(),
                hasAppearanceStreams = proof.FilledRaw.Contains("/AP << /N", StringComparison.Ordinal)
            },
            flattened = new {
                proof.FlattenedInfo.HasForms,
                containsTextAppearance = proof.FlattenedAppearanceText.Contains("<56697369626C652056616C7565> Tj", StringComparison.Ordinal),
                containsChoiceAppearance = proof.FlattenedAppearanceText.Contains("<506F6C616E64> Tj", StringComparison.Ordinal),
                containsCheckBoxAppearance = proof.FlattenedAppearanceText.Contains("1.25 w", StringComparison.Ordinal),
                containsRadioAppearance = proof.FlattenedAppearanceText.Contains("Wire selected", StringComparison.Ordinal),
                containsMultilineAppearance = proof.FlattenedAppearanceText.Contains("<4C696E65206F6E65> Tj", StringComparison.Ordinal) &&
                    proof.FlattenedAppearanceText.Contains("<4C696E652074776F> Tj", StringComparison.Ordinal),
                containsCombAppearance = proof.FlattenedAppearanceText.Contains("<5A> Tj", StringComparison.Ordinal) &&
                    proof.FlattenedAppearanceText.Contains("<58> Tj", StringComparison.Ordinal) &&
                    proof.FlattenedAppearanceText.Contains("<39> Tj", StringComparison.Ordinal) &&
                    proof.FlattenedAppearanceText.Contains("<31> Tj", StringComparison.Ordinal) &&
                    !proof.FlattenedAppearanceText.Contains("<5A583931> Tj", StringComparison.Ordinal),
                containsMultiSelectChoiceAppearance = proof.FlattenedAppearanceText.Contains("<4765726D616E79> Tj", StringComparison.Ordinal) &&
                    proof.FlattenedAppearanceText.Contains("<556E6974656420537461746573> Tj", StringComparison.Ordinal)
            }
        };

        WriteReviewArtifact("pdf-form-appearance-source.pdf", proof.Source);
        WriteReviewArtifact("pdf-form-appearance-filled.pdf", proof.Filled);
        WriteReviewArtifact("pdf-form-appearance-flattened.pdf", proof.Flattened);
        WriteReviewArtifact("pdf-form-appearance-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions { WriteIndented = true }));
    }

    [Fact]
    public void PdfProviderShapedText_ProducesManifestedReviewProof() {
        string? trueTypePath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        string? cffPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        if (trueTypePath == null || cffPath == null) {
            AssertReviewArtifactPrerequisite(
                "The provider-shaped-text review proof requires a local TrueType font and the bundled OpenType/CFF fixture.");
            return;
        }

        const string complexText = "\u0633\u0644\u0627\u0645";
        byte[] trueTypeBytes = File.ReadAllBytes(trueTypePath);
        byte[] cffBytes = File.ReadAllBytes(cffPath);
        PdfCore.PdfTrueTypeFontProgram trueTypeFont = PdfCore.PdfTrueTypeFontProgram.Parse(trueTypeBytes, "OfficeIMO Provider TrueType");
        PdfCore.PdfOpenTypeCffFontProgram cffFont = PdfCore.PdfOpenTypeCffFontProgram.Parse(cffBytes, "OfficeIMO Provider CFF");
        if (PdfCore.PdfTextDiagnostics.AnalyzeEmbeddedFontText(complexText, trueTypeFont).Count > 0 ||
            !TryCreateCffOfficeLigatureGlyphs(cffFont, out IReadOnlyList<DrawingCore.OfficeShapedGlyph>? officeGlyphs)) {
            AssertReviewArtifactPrerequisite(
                "The provider-shaped-text review proof requires TrueType Arabic coverage and the bundled OpenType/CFF office ligature.");
            return;
        }

        var provider = new ManifestTextShapingProvider(
            complexText,
            CreateTrueTypeGlyphMap(complexText, trueTypeFont),
            "office",
            officeGlyphs);
        var report = new PdfCore.PdfConversionReport();
        var options = new PdfCore.PdfOptions {
                CompressContentStreams = false,
                CompressEmbeddedFonts = false
            }
            .ReportDiagnosticsTo(report, "OfficeIMO.Tests")
            .EmbedStandardFont(PdfCore.PdfStandardFont.Helvetica, trueTypeBytes, "OfficeIMO Provider TrueType")
            .EmbedStandardFont(PdfCore.PdfStandardFont.TimesRoman, cffBytes, "OfficeIMO Provider CFF")
            .SetTextShapingProvider(provider);

        byte[] pdf = PdfCore.PdfDocument.Create(options)
            .H1("Provider Shaped Text Gate")
            .Paragraph(paragraph => paragraph.Text(complexText))
            .Paragraph(paragraph => paragraph.Font(PdfCore.PdfStandardFont.TimesRoman).Text("office"))
            .ToBytes();

        string raw = Encoding.ASCII.GetString(pdf);
        string extracted = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.True(provider.TrueTypeCalls >= 1);
        Assert.True(provider.OpenTypeCffCalls >= 1);
        Assert.Contains(complexText, extracted, StringComparison.Ordinal);
        Assert.Contains("office", extracted, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-complex-script-shaping");
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "unsupported-bidirectional-text-layout");
        Assert.Contains("006600660069", raw, StringComparison.Ordinal);
        Assert.DoesNotContain(report.Warnings, warning => warning.Code == "opentype-cff-charstrings-not-subset");
        Assert.True(options.TryGetEmbeddedStandardOpenTypeCffFontProgram(PdfCore.PdfStandardFont.TimesRoman, out PdfCore.PdfOpenTypeCffFontProgram? cffProgram));
        Assert.NotNull(cffProgram);
        PdfCore.PdfOpenTypeCffCompactFontFile cffPlan = cffProgram!.BuildCompactOpenTypeFontFilePlan();
        PdfCore.PdfCffCharStringSubset cffSubset = Assert.IsType<PdfCore.PdfCffCharStringSubset>(cffPlan.CharStringSubset);
        Assert.True(cffPlan.Data.Length < cffBytes.Length);
        Assert.True(cffSubset.IsSubset);
        Assert.True(cffSubset.PrunedGlyphCount > 0);
        Assert.True(cffSubset.SubsetProgramBytes < cffSubset.OriginalProgramBytes);
        Assert.Contains("GSUB", cffPlan.RemovedTables, StringComparer.Ordinal);
        Assert.Contains("GPOS", cffPlan.RemovedLayoutTables, StringComparer.Ordinal);

        var summary = new {
            scenario = "pdf-provider-shaped-text",
            provider.CallCount,
            provider.TrueTypeCalls,
            provider.OpenTypeCffCalls,
            extractedMarkers = new[] { complexText, "office" },
            suppressedWarnings = new[] { "unsupported-complex-script-shaping", "unsupported-bidirectional-text-layout" },
            cffLigatureMappedToUnicode = raw.Contains("006600660069", StringComparison.Ordinal),
            cffCharstringsSubset = cffSubset.IsSubset,
            cffCharstringsRetained = false,
            compactCffEmbedding = new {
                originalFontFileLength = cffBytes.Length,
                embeddedFontFileLength = cffPlan.Data.Length,
                cffTableLength = cffProgram.CffTableLength,
                retainedCffGlyphCount = cffSubset.RetainedGlyphCount,
                usedGlyphCount = cffProgram.GetUsedGlyphIds().Count,
                unusedCffGlyphCount = cffSubset.PrunedGlyphCount,
                originalCharStringBytes = cffSubset.OriginalProgramBytes,
                subsetCharStringBytes = cffSubset.SubsetProgramBytes,
                openTypeTablesEmbedded = cffPlan.EmbeddedTables,
                openTypeTablesRemoved = cffPlan.RemovedTables,
                openTypeLayoutTablesRemoved = cffPlan.RemovedLayoutTables
            },
            warningCodes = report.Warnings.Select(warning => warning.Code).Distinct(StringComparer.Ordinal).ToArray()
        };

        WriteReviewArtifact("pdf-provider-shaped-text.pdf", pdf);
        WriteReviewArtifact("pdf-provider-shaped-text-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions { WriteIndented = true }));
    }

    [Fact]
    public void ExcelDashboardReport_ProducesManifestedReviewProof() {
        byte[] pdf = CreateExcelDashboardReportPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.True(pdf.Length > 0);
        Assert.Contains("Excel Dashboard PDF Gate", text, StringComparison.Ordinal);
        Assert.Contains("KPI Trend", text, StringComparison.Ordinal);
        Assert.Contains("Renewals", text, StringComparison.Ordinal);
        Assert.Contains("Pipeline risk", text, StringComparison.Ordinal);
        Assert.Contains(logical.Images, image => image.Width > 0 && image.Height > 0);

        WriteReviewArtifact("excel-dashboard-report.pdf", pdf);
    }

    [Fact]
    public void PowerPointLayoutThemeGroups_ProducesManifestedReviewProof() {
        byte[] pdf = CreatePowerPointLayoutThemeGroupsPdf();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
        string raw = Encoding.ASCII.GetString(pdf);

        Assert.True(pdf.Length > 0);
        Assert.Contains("Layout Theme Group Gate", text, StringComparison.Ordinal);
        Assert.Contains("Grouped transform marker", text, StringComparison.Ordinal);
        Assert.Contains("20 100 60 40 re", raw, StringComparison.Ordinal);
        Assert.Contains("100 100 60 40 re", raw, StringComparison.Ordinal);
        Assert.Contains("0.122 0.306 0.475 rg", raw, StringComparison.Ordinal);

        WriteReviewArtifact("powerpoint-layout-theme-groups.pdf", pdf);
    }

    [Fact]
    public void PdfLogicalDiagnostics_ProducesManifestedReadbackProof() {
        byte[] pdf = CreateLogicalDiagnosticsPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        string html = PdfHtmlConverterExtensions.ToHtml(logical, options);
        PdfCore.PdfLogicalTableExtraction extraction = Assert.Single(PdfCore.PdfLogicalTableAnalysis.ExtractTables(logical));
        PdfCore.PdfLogicalTableColumnProfile scoreProfile = Assert.Single(extraction.Data.ColumnProfiles, profile => profile.Name == "Score");
        PdfCore.PdfLogicalImage wideImage = Assert.Single(logical.Images, image => image.Width == 3 && image.Height == 2);
        PdfCore.PdfLogicalImage tallImage = Assert.Single(logical.Images, image => image.Width == 2 && image.Height == 3);
        ReaderChunk readerChunk = Assert.Single(PdfReaderAdapter.Read(
            logical,
            sourceName: "pdf-logical-diagnostics-source.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }));

        Assert.Contains(logical.Headings, heading => heading.Text == "Revenue Readback Diagnostics");
        Assert.True(wideImage.HasPlacements);
        Assert.True(tallImage.HasPlacements);
        Assert.Equal(1, wideImage.PlacementCount);
        Assert.Equal(1, tallImage.PlacementCount);
        Assert.True(wideImage.PrimaryPlacement!.IsAxisAligned);
        Assert.True(tallImage.PrimaryPlacement!.IsAxisAligned);
        Assert.Equal(48D, wideImage.PlacedWidth!.Value, 3);
        Assert.Equal(32D, wideImage.PlacedHeight!.Value, 3);
        Assert.Equal(32D, tallImage.PlacedWidth!.Value, 3);
        Assert.Equal(48D, tallImage.PlacedHeight!.Value, 3);
        Assert.Equal(new[] { "Metric", "Score", "Owner" }, extraction.Data.Columns);
        Assert.True(extraction.Data.Diagnostics.Confidence >= 0.95D);
        Assert.Equal(1D, extraction.Data.Diagnostics.SchemaConfidence, 3);
        Assert.Equal(1D, extraction.Data.Diagnostics.CellCompleteness, 3);
        Assert.Equal(1D, extraction.Data.Diagnostics.ColumnGeometryConfidence, 3);
        Assert.Equal(PdfCore.PdfLogicalTableColumnKind.Numeric, scoreProfile.Kind);
        Assert.Equal(1D, scoreProfile.Confidence, 3);
        Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\"", html, StringComparison.Ordinal);
        Assert.Contains("Revenue Readback Diagnostics", html, StringComparison.Ordinal);
        Assert.NotNull(readerChunk.Tables);
        ReaderTable readerTable = Assert.Single(readerChunk.Tables!);
        Assert.NotNull(readerTable.Diagnostics);
        Assert.True(readerTable.Diagnostics!.Confidence >= 0.95D);
        Assert.Equal(1D, readerTable.Diagnostics.ColumnGeometryConfidence, 3);
        Assert.NotNull(readerChunk.Diagnostics);
        Assert.Equal(1, readerChunk.Diagnostics!.TableCount);
        Assert.Equal(1, readerChunk.Diagnostics.TableGeometryCount);
        Assert.Equal(1D, readerChunk.Diagnostics.TableGeometryCoverage, 3);
        Assert.True(readerChunk.Diagnostics.MinTableConfidence >= 0.95D);
        Assert.True(readerChunk.Diagnostics.AverageTableConfidence >= 0.95D);
        Assert.Equal(2, readerChunk.Diagnostics.ImageCount);
        Assert.Equal(2, readerChunk.Diagnostics.ImageGeometryCount);
        Assert.Equal(1D, readerChunk.Diagnostics.ImageGeometryCoverage, 3);
        Assert.NotNull(readerChunk.Visuals);
        Assert.Equal(2, readerChunk.Visuals!.Count);
        ReaderVisual readerWideImage = Assert.Single(readerChunk.Visuals, visual => visual.Width == 3D && visual.Height == 2D);
        ReaderVisual readerTallImage = Assert.Single(readerChunk.Visuals, visual => visual.Width == 2D && visual.Height == 3D);
        Assert.Equal(48D, readerWideImage.PlacedWidth!.Value, 3);
        Assert.Equal(32D, readerWideImage.PlacedHeight!.Value, 3);
        Assert.Equal(32D, readerTallImage.PlacedWidth!.Value, 3);
        Assert.Equal(48D, readerTallImage.PlacedHeight!.Value, 3);
        Assert.All(readerChunk.Visuals, visual => {
            Assert.Equal("image", visual.Kind);
            Assert.Equal("pdf-image", visual.Language);
            Assert.True(visual.HasGeometry);
            Assert.True(visual.IsAxisAligned);
            Assert.NotNull(visual.Location);
            Assert.False(string.IsNullOrWhiteSpace(visual.Location!.BlockAnchor));
        });

        var summary = new {
            scenario = "pdf-logical-diagnostics-readback",
            chunkDiagnostics = new {
                readerChunk.Diagnostics.TableCount,
                readerChunk.Diagnostics.TableGeometryCount,
                readerChunk.Diagnostics.TableGeometryCoverage,
                readerChunk.Diagnostics.MinTableConfidence,
                readerChunk.Diagnostics.AverageTableConfidence,
                readerChunk.Diagnostics.ImageCount,
                readerChunk.Diagnostics.ImageGeometryCount,
                readerChunk.Diagnostics.ImageGeometryCoverage
            },
            table = new {
                readerTable.Kind,
                readerTable.TotalRowCount,
                readerTable.Diagnostics.Confidence,
                readerTable.Diagnostics.SchemaConfidence,
                readerTable.Diagnostics.CellCompleteness,
                readerTable.Diagnostics.ColumnGeometryConfidence,
                readerTable.Diagnostics.Width,
                readerTable.Diagnostics.Height
            },
            visuals = readerChunk.Visuals.Select(visual => new {
                visual.Kind,
                visual.Language,
                visual.SourceName,
                visual.MimeType,
                visual.Width,
                visual.Height,
                visual.X,
                visual.Y,
                visual.PlacedWidth,
                visual.PlacedHeight,
                visual.PlacementCount,
                visual.HasGeometry,
                visual.IsAxisAligned,
                Page = visual.Location?.Page,
                Anchor = visual.Location?.BlockAnchor
            }).ToArray()
        };

        WriteReviewArtifact("pdf-logical-diagnostics-source.pdf", pdf);
        WriteReviewArtifact("pdf-logical-diagnostics-positioned-review.html", Encoding.UTF8.GetBytes(html));
        WriteReviewArtifact("pdf-logical-diagnostics-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions { WriteIndented = true }));
    }

    [Fact]
    public void PdfReaderPageChunks_DoNotRepeatDocumentCatalogActions() {
        byte[] pdf = CreateCatalogActionsMultiPagePdf();

        List<ReaderChunk> pageChunks = PdfReaderAdapter.Read(
            new MemoryStream(pdf, writable: false),
            sourceName: "pdf-catalog-actions-multipage.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList();

        Assert.Equal(2, pageChunks.Count);
        Assert.All(pageChunks, chunk => {
            Assert.DoesNotContain(chunk.Actions ?? Array.Empty<ReaderActionSummary>(), action => action.Scope == ReaderActionScope.Catalog);
            Assert.DoesNotContain(chunk.Actions ?? Array.Empty<ReaderActionSummary>(), action => action.Scope == ReaderActionScope.DocumentOpen);
            Assert.NotNull(chunk.Diagnostics);
            Assert.Equal(2, chunk.Diagnostics!.CatalogActionCount);
            Assert.True(chunk.Diagnostics.HasCatalogActions);
        });

        ReaderChunk documentChunk = Assert.Single(PdfReaderAdapter.Read(
            new MemoryStream(pdf, writable: false),
            sourceName: "pdf-catalog-actions-multipage.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 },
            pdfOptions: new ReaderPdfOptions { ChunkByPage = false }).ToList());

        Assert.NotNull(documentChunk.Actions);
        Assert.Equal(2, documentChunk.Actions!.Count(action => action.Scope == ReaderActionScope.Catalog));
    }

    [Fact]
    public void PdfReaderDegradationCorpus_ProducesManifestedReaderProof() {
        byte[] pdf = CreateReaderDegradationCorpusPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            new MemoryStream(pdf, writable: false),
            sourceName: "pdf-reader-degradation-corpus.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        string text = chunk.Markdown ?? chunk.Text;
        Assert.Contains("Reader Degradation Corpus", text, StringComparison.Ordinal);
        Assert.Contains("Accepted degradation marker", text, StringComparison.Ordinal);
        Assert.Contains("Form and active-content marker", text, StringComparison.Ordinal);
        Assert.DoesNotContain("app.alert", text, StringComparison.Ordinal);
        Assert.Contains(logical.GetLinksByUri("https://example.com/reader-degradation"), link => link.Contents == "Review link");
        Assert.NotNull(chunk.FormFields);
        ReaderFormField field = Assert.Single(chunk.FormFields!, item => item.Name == "Corpus.Contact");
        Assert.Equal(ReaderFormFieldKind.Text, field.Kind);
        Assert.Equal("review@example.com", field.Value);
        Assert.Equal(1, field.WidgetCount);
        Assert.Equal(new[] { 1 }, field.PageNumbers);
        Assert.NotNull(chunk.Diagnostics);
        Assert.Equal(1, chunk.Diagnostics!.LinkCount);
        Assert.Equal(1, chunk.Diagnostics.FormFieldCount);
        Assert.Equal(1, chunk.Diagnostics.SelectedFormWidgetCount);
        Assert.True(chunk.Diagnostics.HasPageActions);
        Assert.True(chunk.Diagnostics.HasAnnotationActions);
        Assert.True(chunk.Diagnostics.HasActiveContent);
        Assert.Equal(1, chunk.Diagnostics.PageActionCount);
        Assert.Equal(1, chunk.Diagnostics.SelectedPageActionCount);
        Assert.Equal(1, chunk.Diagnostics.AnnotationActionCount);
        Assert.Equal(1, chunk.Diagnostics.SelectedAnnotationActionCount);
        Assert.Equal(1, chunk.Diagnostics.PotentiallyUnsafeActionCount);
        Assert.Equal(1, chunk.Diagnostics.JavaScriptActionCount);
        Assert.Equal(0, chunk.Diagnostics.LaunchActionCount);
        Assert.Equal(0, chunk.Diagnostics.SubmitFormActionCount);
        Assert.Equal(0, chunk.Diagnostics.ImportDataActionCount);
        Assert.NotNull(chunk.Actions);
        Assert.Equal(2, chunk.Actions!.Count);
        ReaderActionSummary action = Assert.Single(chunk.Actions, item => item.Scope == ReaderActionScope.Page);
        Assert.Equal(ReaderActionScope.Page, action.Scope);
        Assert.Equal("JavaScript", action.ActionType);
        Assert.Equal("Page/AA", action.Source);
        Assert.Equal("O", action.TriggerName);
        Assert.Equal("O", action.ActionPath);
        Assert.Equal(1, action.PageNumber);
        Assert.False(action.IsChainedAction);
        Assert.True(action.IsPotentiallyUnsafe);
        Assert.DoesNotContain("app.alert", action.ActionType, StringComparison.Ordinal);
        Assert.DoesNotContain("app.alert", action.Source ?? string.Empty, StringComparison.Ordinal);
        Assert.DoesNotContain("app.alert", action.ActionPath ?? string.Empty, StringComparison.Ordinal);
        ReaderActionSummary annotationAction = Assert.Single(chunk.Actions, item => item.Scope == ReaderActionScope.Annotation);
        Assert.Equal("URI", annotationAction.ActionType);
        Assert.Equal("Annotation/A", annotationAction.Source);
        Assert.Equal("Link", annotationAction.Name);
        Assert.Equal("A", annotationAction.ActionPath);
        Assert.Equal(1, annotationAction.PageNumber);
        Assert.False(annotationAction.IsChainedAction);
        Assert.False(annotationAction.IsPotentiallyUnsafe);

        OfficeDocumentReadResult readResult = PdfReaderAdapter.ReadDocument(
            logical,
            sourceName: "pdf-reader-degradation-corpus.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-action-potentially-unsafe-count" && entry.Value == "1");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-action-type-javascript-count" && entry.Value == "1");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-action-type-uri-count" && entry.Value == "1");

        using JsonDocument chunkJson = JsonDocument.Parse(new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Chunks = new[] { chunk }
        }.ToJson());
        JsonElement jsonDiagnostics = chunkJson.RootElement.GetProperty("chunks")[0].GetProperty("diagnostics");
        Assert.Equal(OfficeDocumentReadResultSchema.CurrentVersion,
            chunkJson.RootElement.GetProperty("schemaVersion").GetInt32());
        Assert.Equal(1, jsonDiagnostics.GetProperty("potentiallyUnsafeActionCount").GetInt32());
        Assert.Equal(1, jsonDiagnostics.GetProperty("javaScriptActionCount").GetInt32());
        JsonElement jsonActions = chunkJson.RootElement.GetProperty("chunks")[0].GetProperty("actions");
        Assert.True(jsonActions[0].GetProperty("isPotentiallyUnsafe").GetBoolean());
        Assert.False(jsonActions[1].GetProperty("isPotentiallyUnsafe").GetBoolean());

        var summary = new {
            scenario = "pdf-reader-degradation-corpus",
            acceptedDegradations = new[] {
                "active actions are detected as passive diagnostics only",
                "annotation actions are summarized without executable payloads",
                "script payload text is not emitted into Reader chunk text",
                "form fields are exposed as typed metadata and widget geometry, not editable PDF reconstruction"
            },
            diagnostics = new {
                chunk.Diagnostics.LinkCount,
                chunk.Diagnostics.FormFieldCount,
                chunk.Diagnostics.SelectedFormWidgetCount,
                chunk.Diagnostics.HasPageActions,
                chunk.Diagnostics.HasActiveContent,
                chunk.Diagnostics.PageActionCount,
                chunk.Diagnostics.SelectedPageActionCount,
                chunk.Diagnostics.AnnotationActionCount,
                chunk.Diagnostics.SelectedAnnotationActionCount,
                chunk.Diagnostics.PotentiallyUnsafeActionCount,
                chunk.Diagnostics.JavaScriptActionCount,
                chunk.Diagnostics.LaunchActionCount,
                chunk.Diagnostics.SubmitFormActionCount,
                chunk.Diagnostics.ImportDataActionCount
            },
            actions = chunk.Actions!.Select(item => new {
                Scope = item.Scope.ToString(),
                item.ActionType,
                item.Source,
                item.TriggerName,
                item.ActionPath,
                item.PageNumber,
                item.IsChainedAction,
                item.IsPotentiallyUnsafe
            }).ToArray(),
            formFields = chunk.FormFields!.Select(item => new {
                item.Name,
                Kind = item.Kind.ToString(),
                item.Value,
                item.WidgetCount,
                item.PageNumbers
            }).ToArray()
        };

        WriteReviewArtifact("pdf-reader-degradation-corpus.pdf", pdf);
        WriteReviewArtifact("pdf-reader-degradation-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions {
            WriteIndented = true
        }));
    }

    [Fact]
    public void PdfReaderHostileActionCorpus_ProducesManifestedReaderAndHtmlProof() {
        byte[] pdf = CreateReaderHostileActionCorpusPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        var htmlOptions = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            logical,
            sourceName: "pdf-reader-hostile-action-corpus.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 },
            pdfOptions: new ReaderPdfOptions { ChunkByPage = false }).ToList());
        PdfHtmlConversionResult htmlResult = PdfHtmlConverterExtensions.ToHtmlResult(logical, htmlOptions);

        string text = chunk.Markdown ?? chunk.Text;
        Assert.Contains("Hostile Action Corpus", text, StringComparison.Ordinal);
        Assert.Contains("Passive diagnostics marker", text, StringComparison.Ordinal);
        Assert.DoesNotContain("app.alert", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("tool.exe", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("https://example.com/submit", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("payload.fdf", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("name-tree-data.fdf", text, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("preview.mov", text, StringComparison.OrdinalIgnoreCase);

        Assert.NotNull(chunk.Diagnostics);
        Assert.True(chunk.Diagnostics!.HasOpenAction);
        Assert.True(chunk.Diagnostics.HasCatalogActions);
        Assert.True(chunk.Diagnostics.HasPageActions);
        Assert.True(chunk.Diagnostics.HasAnnotationActions);
        Assert.True(chunk.Diagnostics.HasActiveContent);
        Assert.Equal(5, chunk.Diagnostics.CatalogActionCount);
        Assert.Equal(2, chunk.Diagnostics.PageActionCount);
        Assert.Equal(2, chunk.Diagnostics.SelectedPageActionCount);
        Assert.Equal(3, chunk.Diagnostics.AnnotationActionCount);
        Assert.Equal(3, chunk.Diagnostics.SelectedAnnotationActionCount);
        Assert.Equal(10, chunk.Diagnostics.PotentiallyUnsafeActionCount);
        Assert.Equal(4, chunk.Diagnostics.JavaScriptActionCount);
        Assert.Equal(2, chunk.Diagnostics.LaunchActionCount);
        Assert.Equal(1, chunk.Diagnostics.SubmitFormActionCount);
        Assert.Equal(2, chunk.Diagnostics.ImportDataActionCount);

        Assert.NotNull(chunk.Actions);
        Assert.Equal(11, chunk.Actions!.Count);
        Assert.Contains(chunk.Actions, action => action.Scope == ReaderActionScope.DocumentOpen && action.ActionType == "Destination");
        Assert.Contains(chunk.Actions, action => action.Scope == ReaderActionScope.Catalog && action.ActionType == "JavaScript" && action.Name == "Startup");
        Assert.Contains(chunk.Actions, action => action.Scope == ReaderActionScope.Catalog && action.ActionType == "JavaScript" && action.Name == "Deferred");
        Assert.Contains(chunk.Actions, action => action.Scope == ReaderActionScope.Catalog && action.ActionType == "JavaScript" && action.Name == "AA.WC" && action.TriggerName == "WC" && action.ActionPath == "AA.WC");
        Assert.Contains(chunk.Actions, action => action.Scope == ReaderActionScope.Catalog && action.ActionType == "ImportData" && action.Name == "AA.WC.Next.0" && action.IsChainedAction && action.ActionPath == "AA.WC.Next.0");
        Assert.Contains(chunk.Actions, action => action.Scope == ReaderActionScope.Catalog && action.ActionType == "Movie" && action.Name == "AA.WC.Next.1" && action.IsChainedAction && action.ActionPath == "AA.WC.Next.1");
        Assert.Contains(chunk.Actions, action => action.Scope == ReaderActionScope.Page && action.ActionType == "JavaScript" && action.TriggerName == "O");
        Assert.Contains(chunk.Actions, action => action.Scope == ReaderActionScope.Page && action.ActionType == "Launch" && action.TriggerName == "C");
        Assert.Contains(chunk.Actions, action => action.Scope == ReaderActionScope.Annotation && action.ActionType == "SubmitForm" && action.Source == "Annotation/A");
        Assert.Contains(chunk.Actions, action => action.Scope == ReaderActionScope.Annotation && action.ActionType == "Launch" && action.Source == "Annotation/AA");
        Assert.Contains(chunk.Actions, action => action.Scope == ReaderActionScope.Annotation && action.ActionType == "ImportData" && action.Source == "Annotation/Next" && action.IsChainedAction);
        Assert.Equal(10, chunk.Actions.Count(action => action.IsPotentiallyUnsafe));
        Assert.Equal(3, chunk.Actions.Count(action => action.IsChainedAction));
        Assert.All(chunk.Actions, action => {
            Assert.DoesNotContain("app.alert", action.ActionType, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("tool.exe", action.ActionType, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("https://example.com/submit", action.ActionType, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("payload.fdf", action.ActionType, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("name-tree-data.fdf", action.ActionType, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("preview.mov", action.ActionType, StringComparison.OrdinalIgnoreCase);
        });

        OfficeDocumentReadResult readResult = PdfReaderAdapter.ReadDocument(
            logical,
            sourceName: "pdf-reader-hostile-action-corpus.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-action-count" && entry.Value == "11");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-active-action-count" && entry.Value == "10");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-action-chained-count" && entry.Value == "3");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-action-potentially-unsafe-count" && entry.Value == "10");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-action-type-javascript-count" && entry.Value == "4");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-action-type-launch-count" && entry.Value == "2");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-action-type-submitform-count" && entry.Value == "1");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-action-type-importdata-count" && entry.Value == "2");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-action-type-movie-count" && entry.Value == "1");
        string resultJson = readResult.ToJson();
        Assert.DoesNotContain("app.alert", resultJson, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("tool.exe", resultJson, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("https://example.com/submit", resultJson, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("payload.fdf", resultJson, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("name-tree-data.fdf", resultJson, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("preview.mov", resultJson, StringComparison.OrdinalIgnoreCase);

        Assert.True(htmlResult.Summary.HasOpenAction);
        Assert.True(htmlResult.Summary.HasCatalogActions);
        Assert.True(htmlResult.Summary.HasPageActions);
        Assert.True(htmlResult.Summary.HasAnnotationActions);
        Assert.True(htmlResult.Summary.HasActiveContent);
        Assert.Equal(5, htmlResult.Summary.CatalogActionCount);
        Assert.Equal(2, htmlResult.Summary.PageActionCount);
        Assert.Equal(2, htmlResult.Summary.SelectedPageActionCount);
        Assert.Equal(3, htmlResult.Summary.AnnotationActionCount);
        Assert.Equal(3, htmlResult.Summary.SelectedAnnotationActionCount);
        Assert.Equal(10, htmlResult.Summary.PotentiallyUnsafeActionCount);
        Assert.Equal(4, htmlResult.Summary.JavaScriptActionCount);
        Assert.Equal(2, htmlResult.Summary.LaunchActionCount);
        Assert.Equal(1, htmlResult.Summary.SubmitFormActionCount);
        Assert.Equal(2, htmlResult.Summary.ImportDataActionCount);
        Assert.Contains("Hostile Action Corpus", htmlResult.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("app.alert", htmlResult.Value, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("tool.exe", htmlResult.Value, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("https://example.com/submit", htmlResult.Value, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("payload.fdf", htmlResult.Value, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("name-tree-data.fdf", htmlResult.Value, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("preview.mov", htmlResult.Value, StringComparison.OrdinalIgnoreCase);

        var summary = new {
            scenario = "pdf-reader-hostile-action-corpus",
            acceptedDegradations = new[] {
                "nested catalog JavaScript name trees, catalog additional actions, page actions, and annotation actions are exposed only as passive diagnostics",
                "document-open actions are summarized without executing or embedding action payload text",
                "JavaScript, Launch, SubmitForm, ImportData, and Movie actions are counted as potentially unsafe but remain inert in Reader and HTML output",
                "catalog chained actions carry stable action paths and chained-action flags in Reader summaries"
            },
            readerDiagnostics = new {
                chunk.Diagnostics.HasOpenAction,
                chunk.Diagnostics.HasCatalogActions,
                chunk.Diagnostics.HasPageActions,
                chunk.Diagnostics.HasAnnotationActions,
                chunk.Diagnostics.HasActiveContent,
                chunk.Diagnostics.CatalogActionCount,
                chunk.Diagnostics.PageActionCount,
                chunk.Diagnostics.SelectedPageActionCount,
                chunk.Diagnostics.AnnotationActionCount,
                chunk.Diagnostics.SelectedAnnotationActionCount,
                chunk.Diagnostics.PotentiallyUnsafeActionCount,
                chunk.Diagnostics.JavaScriptActionCount,
                chunk.Diagnostics.LaunchActionCount,
                chunk.Diagnostics.SubmitFormActionCount,
                chunk.Diagnostics.ImportDataActionCount
            },
            htmlSummary = htmlResult.Summary,
            actions = chunk.Actions.Select(item => new {
                Scope = item.Scope.ToString(),
                item.ActionType,
                item.Source,
                item.Name,
                item.TriggerName,
                item.ActionPath,
                item.PageNumber,
                item.IsChainedAction,
                item.IsPotentiallyUnsafe
            }).ToArray(),
            payloadPolicy = "Executable action payloads are not emitted into Reader text, Reader JSON, or positioned review HTML."
        };

        WriteReviewArtifact("pdf-reader-hostile-action-corpus.pdf", pdf);
        WriteReviewArtifact("pdf-reader-hostile-action-positioned-review.html", Encoding.UTF8.GetBytes(htmlResult.Value));
        WriteReviewArtifact("pdf-reader-hostile-action-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions {
            WriteIndented = true
        }));
    }

    [Fact]
    public void PdfReaderHostileLayoutCorpus_ProducesManifestedReaderProof() {
        byte[] pdf = CreateReaderHostileLayoutCorpusPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = false
        });

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            new MemoryStream(pdf, writable: false),
            sourceName: "pdf-reader-hostile-layout-corpus.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        string text = chunk.Markdown ?? chunk.Text;
        Assert.Contains("Hostile Layout Corpus", text, StringComparison.Ordinal);
        Assert.Contains("Left column marker", text, StringComparison.Ordinal);
        Assert.Contains("Right column marker", text, StringComparison.Ordinal);
        Assert.Contains("Rotated note marker", text, StringComparison.Ordinal);
        Assert.Contains(logical.Pages[0].TextBlocks, block => block.Text == "Rotated note marker");
        Assert.NotNull(chunk.Diagnostics);
        Assert.Equal(1, chunk.Diagnostics!.ImageCount);
        Assert.Equal(1, chunk.Diagnostics.ImageGeometryCount);
        Assert.Equal(1D, chunk.Diagnostics.ImageGeometryCoverage, 3);
        Assert.Equal(0, chunk.Diagnostics.TableCount);
        Assert.Null(chunk.Diagnostics.MinTableConfidence);
        Assert.NotNull(chunk.Visuals);
        ReaderVisual visual = Assert.Single(chunk.Visuals!);
        Assert.Equal("image", visual.Kind);
        Assert.True(visual.HasGeometry);
        Assert.Equal(false, visual.IsAxisAligned);
        Assert.Equal(1, visual.PlacementCount);
        Assert.True(visual.PlacedWidth > 0);
        Assert.True(visual.PlacedHeight > 0);

        var summary = new {
            scenario = "pdf-reader-hostile-layout-corpus",
            acceptedDegradations = new[] {
                "close columns and rotated text are exposed as born-digital text, but the Reader contract does not promise perfect human reading order for hostile layouts",
                "skewed image placement is preserved as geometry with IsAxisAligned=false rather than reconstructed into editable Office drawing transforms",
                "no table is emitted because the fixture intentionally lacks stable table ruling or column/header structure"
            },
            diagnostics = new {
                chunk.Diagnostics.PageCount,
                chunk.Diagnostics.SelectedPageCount,
                chunk.Diagnostics.ImageCount,
                chunk.Diagnostics.ImageGeometryCount,
                chunk.Diagnostics.ImageGeometryCoverage,
                chunk.Diagnostics.TableCount,
                chunk.Diagnostics.TableGeometryCoverage,
                chunk.Diagnostics.MinTableConfidence,
                chunk.Diagnostics.AverageTableConfidence
            },
            visuals = chunk.Visuals!.Select(item => new {
                item.Kind,
                item.SourceName,
                item.Width,
                item.Height,
                item.X,
                item.Y,
                item.PlacedWidth,
                item.PlacedHeight,
                item.PlacementCount,
                item.HasGeometry,
                item.IsAxisAligned
            }).ToArray(),
            textMarkers = new[] {
                "Hostile Layout Corpus",
                "Left column marker",
                "Right column marker",
                "Rotated note marker"
            }
        };

        WriteReviewArtifact("pdf-reader-hostile-layout-corpus.pdf", pdf);
        WriteReviewArtifact("pdf-reader-hostile-layout-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions {
            WriteIndented = true
        }));
    }

    [Fact]
    public void PdfReaderHostileTableCorpus_ProducesManifestedReaderProof() {
        byte[] pdf = CreateReaderHostileTableCorpusPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        IReadOnlyList<PdfCore.PdfLogicalTableExtraction> extractions = PdfCore.PdfLogicalTableAnalysis.ExtractTables(logical);

        ReaderChunk chunk = Assert.Single(PdfReaderAdapter.Read(
            new MemoryStream(pdf, writable: false),
            sourceName: "pdf-reader-hostile-table-corpus.pdf",
            pdfOptions: new ReaderPdfOptions {
                LayoutOptions = new PdfCore.PdfTextLayoutOptions {
                    ForceSingleColumn = true
                }
            },
            readerOptions: new ReaderOptions { MaxChars = 8_000 }).ToList());

        string text = chunk.Markdown ?? chunk.Text;
        Assert.Contains("Hostile Table Corpus", text, StringComparison.Ordinal);
        Assert.Contains("Jittered table marker", text, StringComparison.Ordinal);
        Assert.Contains("Alpha team", text, StringComparison.Ordinal);
        Assert.Contains("Gamma team", text, StringComparison.Ordinal);

        PdfCore.PdfLogicalTableExtraction extraction = Assert.Single(extractions);
        Assert.Equal(new[] { "Column 1", "Column 2", "Column 3" }, extraction.Data.Columns);
        Assert.Equal(3, extraction.Data.Rows.Count);
        Assert.Contains(extraction.Data.Rows, row => row[0] == "Alpha team" && row[2] == "72");
        Assert.Contains(extraction.Data.Rows, row => row[0] == "Gamma team" && row[2] == "91");
        Assert.True(extraction.Data.Diagnostics.HasGeometry);
        Assert.True(extraction.Data.Diagnostics.Confidence >= 0.80D);
        Assert.True(extraction.Data.Diagnostics.Confidence < 0.95D);
        Assert.Equal(0.65D, extraction.Data.Diagnostics.SchemaConfidence, 3);
        Assert.Equal(1D, extraction.Data.Diagnostics.CellCompleteness, 3);
        Assert.Equal(1D, extraction.Data.Diagnostics.ColumnGeometryConfidence, 3);

        Assert.NotNull(chunk.Diagnostics);
        Assert.Equal(1, chunk.Diagnostics!.TableCount);
        Assert.Equal(1, chunk.Diagnostics.TableGeometryCount);
        Assert.Equal(1D, chunk.Diagnostics.TableGeometryCoverage, 3);
        Assert.True(chunk.Diagnostics.MinTableConfidence >= 0.80D);
        Assert.True(chunk.Diagnostics.MinTableConfidence < 0.95D);
        Assert.True(chunk.Diagnostics.AverageTableConfidence >= 0.80D);
        Assert.True(chunk.Diagnostics.AverageTableConfidence < 0.95D);
        Assert.Equal(1, chunk.Diagnostics.LowConfidenceTableCount);
        Assert.Equal(1, chunk.Diagnostics.NumericTableColumnCount);
        Assert.Equal(3, chunk.Diagnostics.FallbackTableColumnNameCount);
        Assert.Equal(0, chunk.Diagnostics.MissingTableCellCount);
        Assert.NotNull(chunk.Tables);
        ReaderTable table = Assert.Single(chunk.Tables!);
        Assert.Equal(new[] { "Column 1", "Column 2", "Column 3" }, table.Columns);
        Assert.NotNull(table.Diagnostics);
        Assert.True(table.Diagnostics!.Confidence >= 0.80D);
        Assert.True(table.Diagnostics.Confidence < 0.95D);
        Assert.Equal(ReaderTableColumnKind.Numeric, table.ColumnProfiles[2].Kind);

        OfficeDocumentReadResult readResult = PdfReaderAdapter.ReadDocument(
            logical,
            sourceName: "pdf-reader-hostile-table-corpus.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-table-count" && entry.Value == "1");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-table-low-confidence-count" && entry.Value == "1");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-table-numeric-column-count" && entry.Value == "1");
        Assert.Contains(readResult.Metadata, entry => entry.Id == "pdf-table-fallback-column-name-count" && entry.Value == "3");

        using JsonDocument chunkJson = JsonDocument.Parse(new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Chunks = new[] { chunk }
        }.ToJson());
        JsonElement jsonDiagnostics = chunkJson.RootElement.GetProperty("chunks")[0].GetProperty("diagnostics");
        Assert.Equal(1, jsonDiagnostics.GetProperty("lowConfidenceTableCount").GetInt32());
        Assert.Equal(1, jsonDiagnostics.GetProperty("numericTableColumnCount").GetInt32());
        Assert.Equal(3, jsonDiagnostics.GetProperty("fallbackTableColumnNameCount").GetInt32());
        Assert.Equal(0, jsonDiagnostics.GetProperty("missingTableCellCount").GetInt32());

        var summary = new {
            scenario = "pdf-reader-hostile-table-corpus",
            acceptedDegradations = new[] {
                "headerless table-like bands are emitted with fallback column names",
                "jittered column positions are accepted as best-effort geometry when confidence remains below perfect-table proof thresholds",
                "the Reader contract exposes table confidence and numeric-column hints but does not reconstruct an editable spreadsheet"
            },
            diagnostics = new {
                chunk.Diagnostics.TableCount,
                chunk.Diagnostics.TableGeometryCount,
                chunk.Diagnostics.TableGeometryCoverage,
                chunk.Diagnostics.MinTableConfidence,
                chunk.Diagnostics.AverageTableConfidence,
                chunk.Diagnostics.LowConfidenceTableCount,
                chunk.Diagnostics.NumericTableColumnCount,
                chunk.Diagnostics.FallbackTableColumnNameCount,
                chunk.Diagnostics.MissingTableCellCount
            },
            table = new {
                table.Kind,
                table.Columns,
                table.TotalRowCount,
                table.Diagnostics.Confidence,
                table.Diagnostics.SchemaConfidence,
                table.Diagnostics.CellCompleteness,
                table.Diagnostics.ColumnGeometryConfidence,
                table.Diagnostics.Width,
                table.Diagnostics.Height,
                numericColumns = table.ColumnProfiles
                    .Where(profile => profile.IsNumeric)
                    .Select(profile => new {
                        profile.Index,
                        profile.Name,
                        Kind = profile.Kind.ToString(),
                        profile.Confidence
                    })
                    .ToArray(),
                rows = table.Rows
            }
        };

        WriteReviewArtifact("pdf-reader-hostile-table-corpus.pdf", pdf);
        WriteReviewArtifact("pdf-reader-hostile-table-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions {
            WriteIndented = true
        }));
    }

    [Fact]
    public void PdfReaderOcrHandoffCorpus_ProducesManifestedReaderProof() {
        byte[] pdf = CreateReaderOcrHandoffCorpusPdf();
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(pdf, writable: false),
            sourceName: "pdf-reader-ocr-handoff-corpus.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        OfficeDocumentOcrCandidate candidate = Assert.Single(result.OcrCandidates);
        OfficeDocumentAsset asset = Assert.Single(result.Assets);
        OfficeDocumentDiagnostic diagnostic = Assert.Single(result.Diagnostics, item => item.Code == "ocr-needed");
        OfficeDocumentPage page = Assert.Single(result.Pages);
        OfficeDocumentOcrEnrichmentResult enrichment = result.ApplyOcrResults(new[] {
            new OfficeDocumentOcrTextResult {
                CandidateId = candidate.Id,
                Text = "Scanned statement\nAmount due 123.45 EUR",
                Confidence = 0.96D,
                Language = "en",
                Provider = "external-ocr-contract",
                Model = "fixture"
            }
        });
        OfficeDocumentReadResult enriched = enrichment.Document;

        Assert.Equal("image", candidate.Kind);
        Assert.Equal(1, candidate.Location.Page);
        Assert.Equal(1, candidate.ImageCount);
        Assert.Equal(0, candidate.TextBlockCount);
        Assert.Equal(asset.Id, candidate.AssetId);
        Assert.NotNull(candidate.Region);
        Assert.True(candidate.Region!.Width > 0D);
        Assert.True(candidate.Region.Height > 0D);
        Assert.NotNull(asset.Region);
        Assert.True(asset.Region!.Width > 0D);
        Assert.True(asset.Region.Height > 0D);
        Assert.Same(candidate, Assert.Single(page.OcrCandidates));
        Assert.Equal(1, diagnostic.Location?.Page);
        Assert.Contains("no meaningful native text", diagnostic.Message, StringComparison.Ordinal);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-image-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-ocr-candidate-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-ocr-image-candidate-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-ocr-asset-linked-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-ocr-candidate-geometry-count").Value);
        Assert.Equal("1", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-ocr-candidate-geometry-coverage").Value);
        Assert.Equal(1, enrichment.Report.AppliedResultCount);
        Assert.Equal(0, enrichment.Report.UnresolvedCandidateCount);
        Assert.Empty(enriched.OcrCandidates);
        Assert.DoesNotContain(enriched.Diagnostics, item => item.Code == "ocr-needed");
        Assert.Contains("officeimo.reader.ocr-enrichment", enriched.CapabilitiesUsed);
        OfficeDocumentBlock ocrBlock = Assert.Single(enriched.Blocks, item => item.Kind == "ocr-text");
        Assert.Equal("Scanned statement\nAmount due 123.45 EUR", ocrBlock.Text);
        Assert.Equal(candidate.Location.Page, ocrBlock.Location.Page);
        Assert.Equal("1", Assert.Single(enriched.Metadata, metadata => metadata.Id == "reader-ocr-applied-count").Value);
        Assert.Equal("0", Assert.Single(enriched.Metadata, metadata => metadata.Id == "reader-ocr-unresolved-candidate-count").Value);

        var metadata = result.Metadata
            .Where(entry => entry.Id.StartsWith("pdf-ocr-", StringComparison.Ordinal) || entry.Id.StartsWith("pdf-image-", StringComparison.Ordinal))
            .Select(entry => new {
                entry.Id,
                entry.Category,
                entry.Name,
                entry.Value,
                entry.ValueType
            })
            .ToArray();
        var summary = new {
            scenario = "pdf-reader-ocr-handoff-corpus",
            acceptedDegradations = new[] {
                "image-only PDF pages are surfaced as OCR candidates rather than treated as successfully extracted native text",
                "OfficeIMO.Reader.Pdf reports candidate geometry, linked image asset metadata, and ocr-needed diagnostics but does not run OCR inside the dependency-light core",
                "callers can route the stable read-result OCR candidate contract to an external OCR service or UI review workflow",
                "external OCR text can be merged back through OfficeIMO.Reader as generic ocr-text blocks/chunks with traceable reader.ocr metadata"
            },
            readResult = new {
                result.SchemaId,
                result.SchemaVersion,
                pageCount = result.Pages.Count,
                chunkCount = result.Chunks.Count,
                assetCount = result.Assets.Count,
                ocrCandidateCount = result.OcrCandidates.Count,
                diagnosticCodes = result.Diagnostics.Select(item => item.Code).ToArray()
            },
            candidate = new {
                candidate.Id,
                candidate.Kind,
                candidate.Reason,
                candidate.Confidence,
                candidate.AssetId,
                candidate.ImageCount,
                candidate.TextBlockCount,
                candidate.Location.Page,
                candidate.Region.X,
                candidate.Region.Y,
                candidate.Region.Width,
                candidate.Region.Height
            },
            asset = new {
                asset.Id,
                asset.Kind,
                asset.MediaType,
                asset.Extension,
                asset.FileName,
                asset.Width,
                asset.Height,
                asset.LengthBytes,
                asset.PayloadHash,
                placedX = asset.Region!.X,
                placedY = asset.Region.Y,
                placedWidth = asset.Region.Width,
                placedHeight = asset.Region.Height
            },
            enrichment = new {
                enrichment.Report.CandidateCount,
                enrichment.Report.ResultCount,
                enrichment.Report.AppliedResultCount,
                enrichment.Report.UnresolvedCandidateCount,
                enrichment.Report.UnmatchedResultCount,
                enrichment.Report.EnrichedBlockCount,
                enrichment.Report.EnrichedChunkCount,
                enrichment.Report.AppliedCandidateIds,
                enrichedCandidateCount = enriched.OcrCandidates.Count,
                diagnosticCodes = enriched.Diagnostics.Select(item => item.Code).ToArray(),
                capabilities = enriched.CapabilitiesUsed,
                block = new {
                    ocrBlock.Id,
                    ocrBlock.Kind,
                    ocrBlock.Text,
                    ocrBlock.Location.Page,
                    ocrBlock.Region!.X,
                    ocrBlock.Region.Y,
                    ocrBlock.Region.Width,
                    ocrBlock.Region.Height
                },
                metadata = enriched.Metadata
                    .Where(entry => entry.Id.StartsWith("reader-ocr-", StringComparison.Ordinal))
                    .Select(entry => new {
                        entry.Id,
                        entry.Category,
                        entry.Name,
                        entry.Value,
                        entry.ValueType,
                        entry.SourceObjectId,
                        entry.Attributes
                    })
                    .ToArray()
            },
            metadata
        };

        WriteReviewArtifact("pdf-reader-ocr-handoff-corpus.pdf", pdf);
        WriteReviewArtifact("pdf-reader-ocr-handoff-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions {
            WriteIndented = true
        }));
    }

    [Fact]
    public void PdfReaderXfaFormCorpus_ProducesManifestedReaderProof() {
        byte[] pdf = CreateReaderXfaFormCorpusPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(pdf);
        OfficeDocumentReadResult result = PdfReaderAdapter.ReadDocument(
            new MemoryStream(pdf, writable: false),
            sourceName: "pdf-reader-xfa-form-corpus.pdf",
            readerOptions: new ReaderOptions { MaxChars = 8_000 });

        Assert.NotNull(logical.AcroFormXfa);
        Assert.True(logical.HasAcroFormXfa);
        Assert.Equal("array", logical.AcroFormXfa!.ObjectKind);
        Assert.Equal(2, logical.AcroFormXfa.PacketCount);
        Assert.Equal(new[] { "template", "datasets" }, logical.AcroFormXfa.PacketNames);
        Assert.Equal(2, logical.AcroFormXfa.StreamCount);
        Assert.True(logical.AcroFormXfa.TotalPayloadBytes > 0);
        Assert.True(logical.AcroFormXfa.HasTemplatePacket);
        Assert.True(logical.AcroFormXfa.HasDatasetsPacket);
        Assert.True(info.HasForms);
        Assert.True(info.HasAcroFormXfa);
        Assert.Equal("true", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-acroform-xfa-present").Value);
        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-acroform-xfa-packet-count").Value);
        Assert.Equal("2", Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-acroform-xfa-stream-count").Value);
        OfficeDocumentMetadataEntry xfa = Assert.Single(result.Metadata, metadata => metadata.Id == "pdf-acroform-xfa");
        Assert.Equal("array", xfa.Value);
        Assert.Equal("template,datasets", xfa.Attributes["packetNames"]);
        Assert.Equal("true", xfa.Attributes["hasTemplatePacket"]);
        Assert.Equal("true", xfa.Attributes["hasDatasetsPacket"]);

        var metadata = result.Metadata
            .Where(entry => entry.Id.StartsWith("pdf-acroform-xfa", StringComparison.Ordinal))
            .Select(entry => new {
                entry.Id,
                entry.Category,
                entry.Name,
                entry.Value,
                entry.ValueType,
                entry.Attributes
            })
            .ToArray();
        var summary = new {
            scenario = "pdf-reader-xfa-form-corpus",
            acceptedDegradations = new[] {
                "AcroForm XFA packets are detected and reported as metadata without claiming XFA rendering or filling",
                "OfficeIMO.Pdf exposes XFA packet names, payload counts, and template/datasets flags through the logical model and inspector",
                "OfficeIMO.Reader.Pdf carries the same XFA facts in stable reader metadata for downstream routing"
            },
            logical = new {
                logical.HasAcroFormXfa,
                logical.AcroFormXfa.ObjectKind,
                logical.AcroFormXfa.PacketCount,
                logical.AcroFormXfa.PacketNames,
                logical.AcroFormXfa.StreamCount,
                logical.AcroFormXfa.StringCount,
                logical.AcroFormXfa.DictionaryCount,
                logical.AcroFormXfa.TotalPayloadBytes,
                logical.AcroFormXfa.HasTemplatePacket,
                logical.AcroFormXfa.HasDatasetsPacket
            },
            inspector = new {
                info.HasForms,
                info.HasAcroFormXfa,
                info.FormFieldCount,
                info.AcroFormXfa!.PacketCount
            },
            readResult = new {
                result.SchemaId,
                result.SchemaVersion,
                pageCount = result.Pages.Count,
                formCount = result.Forms.Count,
                metadataCount = metadata.Length
            },
            metadata
        };

        WriteReviewArtifact("pdf-reader-xfa-form-corpus.pdf", pdf);
        WriteReviewArtifact("pdf-reader-xfa-form-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions {
            WriteIndented = true
        }));
    }

    [Fact]
    public void HtmlCssResourcePolicy_ProducesManifestedReviewProof() {
        const string stylesheetUri = "https://allowed.example.test/policy.css";
        var options = new HtmlPdfSaveOptions {
            MaxResourceBytes = 8192,
            MaxTotalResourceBytes = 16384,
            ResourceResolver = (request, cancellationToken) => Task.FromResult<HtmlResolvedResource?>(
                request.Uri.AbsoluteUri == stylesheetUri
                    ? new HtmlResolvedResource(Encoding.UTF8.GetBytes("p.policy-note { color:#123456; }"), "text/css")
                    : null)
        };

        PdfCore.PdfDocumentConversionResult result = HtmlConversionDocument.Parse(CreateCssResourcePolicyHtml(stylesheetUri)).ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
        HtmlPdfResourcePolicySummary policy = options.GetResourcePolicySummary();

        Assert.True(policy.HasResourceResolver);
        Assert.Equal(8192, policy.MaxResourceBytes);
        Assert.Equal(16384, policy.MaxTotalResourceBytes);
        Assert.Contains(result.Report.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ExternalStylesheetPending);
        Assert.Contains("HTML CSS Resource Policy Gate", text, StringComparison.Ordinal);
        Assert.Contains("Local stylesheet marker", text, StringComparison.Ordinal);
        Assert.Contains("Blocked remote stylesheet marker", text, StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");

        var summary = new {
            scenario = "html-css-resource-policy",
            renderer = "direct-html",
            policy,
            diagnostics = result.Report.Warnings
        };

        WriteReviewArtifact("html-css-resource-policy.pdf", pdf);
        WriteReviewArtifact("html-css-resource-policy-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions { WriteIndented = true }));
    }

    [Fact]
    public void Html_ToPdfResult_ReturnsPdfDocumentAndReportSnapshot() {
        const string stylesheetUri = "https://allowed.example.test/policy.css";
        var options = new HtmlPdfSaveOptions {
            ResourceResolver = (request, cancellationToken) => Task.FromResult<HtmlResolvedResource?>(
                request.Uri.AbsoluteUri == stylesheetUri
                    ? new HtmlResolvedResource(Encoding.UTF8.GetBytes("p.policy-note { color:#123456; }"), "text/css")
                    : null)
        };

        PdfCore.PdfDocumentConversionResult result = HtmlConversionDocument.Parse(CreateCssResourcePolicyHtml(stylesheetUri)).ToPdfDocumentResult(options);
        PdfCore.PdfDocumentConversionResult processed = result.Process(document => document.AppendMetadataRevision(title: "Processed HTML PDF"));

        Assert.True(result.HasWarnings);
        Assert.True(processed.HasWarnings);
        Assert.Equal("Processed HTML PDF", processed.Value.Inspect().Metadata.Title);
        Assert.Contains(result.Warnings, warning => warning.Code == HtmlRenderDiagnosticCodes.ExternalStylesheetPending);
        Assert.Contains("HTML CSS Resource Policy Gate", result.Value.Read.Text(), StringComparison.Ordinal);
        Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(result.ToBytes()), image => image.IsImageFile && image.MimeType == "image/png");
    }

    [Fact]
    public void PdfStandardSecurityRoundtrip_ProducesManifestedReviewProof() {
        byte[] encrypted = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions().SetEncryption("open", "owner"))
            .Meta(title: "PDF Standard Security Gate", author: "OfficeIMO")
            .H1("PDF Standard Security Gate")
            .Paragraph(paragraph => paragraph.Text("Credential protected marker"))
            .Paragraph(paragraph => paragraph.Text("Extracted page marker"))
            .ToBytes();

        PdfCore.PdfDocumentPreflight blockedPreflight = PdfCore.PdfInspector.Preflight(encrypted);
        var readOptions = new PdfCore.PdfReadOptions { Password = "open" };
        PdfCore.PdfDocumentPreflight readablePreflight = PdfCore.PdfInspector.Preflight(encrypted, readOptions);
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(encrypted, readOptions);
        string text = PdfCore.PdfTextExtractor.ExtractAllText(encrypted, (PdfCore.PdfTextLayoutOptions?)null, readOptions);
        PdfCore.PdfDocument opened = PdfCore.PdfDocument.Open(encrypted, readOptions);
        IReadOnlyList<PdfCore.PdfDocument> splitPages = opened.Pages.Split();
        byte[] extractedPage = Assert.Single(splitPages).ToBytes();

        Assert.True(PdfCore.PdfInspector.Probe(encrypted).HasEncryption);
        Assert.False(blockedPreflight.CanRead);
        Assert.Contains(blockedPreflight.ReadBlockers, blocker => blocker.Kind == PdfCore.PdfReadBlockerKind.Encryption);
        Assert.Throws<PdfCore.PdfPasswordRequiredException>(() => PdfCore.PdfReadDocument.Open(encrypted));
        Assert.Throws<PdfCore.PdfInvalidPasswordException>(() => PdfCore.PdfReadDocument.Open(encrypted, new PdfCore.PdfReadOptions { Password = "wrong" }));
        Assert.True(readablePreflight.CanRead);
        Assert.False(readablePreflight.CanRewrite);
        Assert.True(info.Security.HasEncryption);
        Assert.Equal("Standard", info.Security.EncryptionFilter);
        Assert.Equal(6, info.Security.EncryptionRevision);
        Assert.Equal(256, info.Security.EncryptionLengthBits);
        Assert.Contains("Credential protected marker", text, StringComparison.Ordinal);
        Assert.Contains("Credential protected marker", opened.Read.Text(), StringComparison.Ordinal);
        Assert.False(PdfCore.PdfInspector.Probe(extractedPage).HasEncryption);
        Assert.Contains("Extracted page marker", PdfCore.PdfTextExtractor.ExtractAllText(extractedPage), StringComparison.Ordinal);
        Assert.Contains(readablePreflight.RewriteBlockers, blocker => blocker.Kind == PdfCore.PdfRewriteBlockerKind.Encryption);

        var summary = new {
            scenario = "pdf-standard-security-roundtrip",
            blockedWithoutPassword = new {
                blockedPreflight.CanRead,
                readBlockers = blockedPreflight.ReadBlockers.Select(blocker => blocker.Kind.ToString()).ToArray()
            },
            openedWithPassword = new {
                readablePreflight.CanRead,
                readablePreflight.CanRewrite,
                info.Security.HasEncryption,
                info.Security.EncryptionFilter,
                info.Security.EncryptionRevision,
                info.Security.EncryptionLengthBits,
                textMarkers = new[] { "Credential protected marker", "Extracted page marker" }
            },
            extractedPage = new {
                hasEncryption = PdfCore.PdfInspector.Probe(extractedPage).HasEncryption,
                byteLength = extractedPage.Length
            },
            acceptedLimit = "OfficeIMO.Pdf supports Standard security password read/decrypt and password-backed page extraction; encrypted rewrite and form mutation still fail closed."
        };

        WriteReviewArtifact("pdf-standard-security-roundtrip.pdf", encrypted);
        WriteReviewArtifact("pdf-standard-security-extracted-page.pdf", extractedPage);
        WriteReviewArtifact("pdf-standard-security-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions { WriteIndented = true }));
    }

    [Fact]
    public void HtmlPdfRoundTripProfiles_ProduceManifestedReviewProof() {
        const string linkUri = "https://example.com/html-pdf-roundtrip";
        var htmlOptions = new HtmlPdfSaveOptions();
        PdfCore.PdfDocumentConversionResult htmlResult = HtmlConversionDocument.Parse(CreatePracticalHtmlSample(linkUri)).ToPdfDocumentResult(htmlOptions);
        byte[] pdf = htmlResult.ToBytes();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        var semanticOptions = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            IncludeLinkAnnotations = true
        };
        var positionedOptions = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        PdfHtmlConversionResult semantic = PdfHtmlConverterExtensions.ToHtmlResult(logical, semanticOptions);
        PdfHtmlConversionResult positioned = PdfHtmlConverterExtensions.ToHtmlResult(logical, positionedOptions);

        Assert.True(pdf.Length > 0);
        Assert.True(logical.PageCount >= 2);
        Assert.Contains(logical.TextBlocks, block => block.Text.Contains("Practical HTML", StringComparison.Ordinal));
        Assert.Contains(logical.GetLinksByUri(linkUri), link => link.Contents == "Report link");
        Assert.Contains(logical.Images, image => image.PlacementCount > 0);
        Assert.Contains("Practical HTML", semantic.Value, StringComparison.Ordinal);
        Assert.Contains("Report link", semantic.Value, StringComparison.Ordinal);
        Assert.Contains("rel=\"noopener noreferrer\"", semantic.Value, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-page\" id=\"pdf-page-1\" data-page-number=\"1\"", positioned.Value, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-image-placeholder\"", positioned.Value, StringComparison.Ordinal);
        Assert.Contains("rel=\"noopener noreferrer\"", positioned.Value, StringComparison.Ordinal);
        Assert.Equal(PdfHtmlProfile.Semantic, semantic.Summary.Profile);
        Assert.Equal(PdfHtmlProfile.PositionedReview, positioned.Summary.Profile);
        Assert.True(semantic.Summary.RenderedPageCount >= 2);
        Assert.True(positioned.Summary.RenderedPageCount >= 2);
        Assert.True(positioned.Summary.TextBlockCount > 0);
        Assert.True(positioned.Summary.ImagePlacementCount > 0);
        Assert.True(semantic.Summary.ImagePlaceholderCount > 0);
        Assert.True(positioned.Summary.ImagePlaceholderCount > 0);
        Assert.True(positioned.Summary.LinkCount > 0);
        Assert.True(semantic.Summary.RenderedLinkCount > 0);
        Assert.True(positioned.Summary.RenderedLinkCount > 0);
        Assert.Equal(semantic.Summary.RenderedLinkCount, semantic.Summary.RenderedSafeUriLinkCount);
        Assert.Equal(positioned.Summary.RenderedLinkCount, positioned.Summary.RenderedSafeUriLinkCount);
        Assert.Equal(0, semantic.Summary.RenderedUnsafeUriLinkCount);
        Assert.Equal(0, positioned.Summary.RenderedUnsafeUriLinkCount);
        Assert.Equal(0, semantic.Summary.SkippedLinkCount);
        Assert.Equal(0, positioned.Summary.SkippedLinkCount);
        Assert.False(semantic.Report.HasWarnings);
        Assert.False(positioned.Report.HasWarnings);

        var summary = new {
            scenario = "html-pdf-roundtrip-profile-contract",
            htmlToPdfRenderer = "direct-html",
            pdfToSemanticProfile = PdfHtmlProfileContracts.Get(PdfHtmlProfile.Semantic),
            pdfToPositionedProfile = PdfHtmlProfileContracts.Get(PdfHtmlProfile.PositionedReview),
            htmlToPdfWarnings = htmlResult.Report.Warnings.Select(warning => new {
                warning.Converter,
                warning.Code,
                warning.Source,
                warning.Message,
                Severity = warning.Severity.ToString(),
                warning.Details
            }).ToArray(),
            semantic = semantic.Summary,
            positioned = positioned.Summary
        };

        WriteReviewArtifact("html-pdf-roundtrip-source.pdf", pdf);
        WriteReviewArtifact("html-pdf-roundtrip-semantic.html", Encoding.UTF8.GetBytes(semantic.Value));
        WriteReviewArtifact("html-pdf-roundtrip-positioned.html", Encoding.UTF8.GetBytes(positioned.Value));
        WriteReviewArtifact("html-pdf-roundtrip-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions {
            WriteIndented = true
        }));
    }

    [Fact]
    public void PdfToHtmlResult_PreservesUnsafeLinksAsInertReviewMetadata() {
        byte[] pdf = CreateLinkAnnotationPdf("javascript:alert(1)");
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf);
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        PdfHtmlConversionResult result = PdfHtmlConverterExtensions.ToHtmlResult(logical, options);

        Assert.Contains("class=\"pdf-link\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-unsafe-href=\"javascript:alert(1)\"", result.Value, StringComparison.Ordinal);
        Assert.DoesNotContain(" href=\"javascript:alert(1)\"", result.Value, StringComparison.Ordinal);
        Assert.Equal(1, result.Summary.LinkCount);
        Assert.Equal(1, result.Summary.RenderedLinkCount);
        Assert.Equal(0, result.Summary.RenderedSafeUriLinkCount);
        Assert.Equal(1, result.Summary.RenderedUnsafeUriLinkCount);
        Assert.Equal(0, result.Summary.RenderedInternalDestinationLinkCount);
        Assert.Equal(0, result.Summary.SkippedLinkCount);
        Assert.False(result.Report.HasWarnings);
    }

    [Fact]
    public void PdfToHtmlResult_PreservesDirectDestinationLinksAsReviewMetadata() {
        byte[] pdf = CreateDirectDestinationLinkPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf);
        var options = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true
        };

        PdfHtmlConversionResult result = PdfHtmlConverterExtensions.ToHtmlResult(logical, options);

        Assert.Contains(logical.GetLinksByDestinationPageNumber(2), link => link.Contents == "Jump to page two");
        Assert.Contains("data-destination-page-number=\"2\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-destination-mode=\"FitRectangle\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-destination-left=\"10\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-destination-bottom=\"20\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-destination-right=\"90\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-destination-top=\"144\"", result.Value, StringComparison.Ordinal);
        Assert.Contains(">Jump to page two</a>", result.Value, StringComparison.Ordinal);
        Assert.Equal(1, result.Summary.LinkCount);
        Assert.Equal(1, result.Summary.RenderedLinkCount);
        Assert.Equal(0, result.Summary.RenderedSafeUriLinkCount);
        Assert.Equal(0, result.Summary.RenderedUnsafeUriLinkCount);
        Assert.Equal(1, result.Summary.RenderedInternalDestinationLinkCount);
        Assert.Equal(0, result.Summary.SkippedLinkCount);
        Assert.False(result.Report.HasWarnings);
    }

    [Fact]
    public void PdfEditableOfficeProfiles_ProduceManifestedProof() {
        byte[] pdf = CreateLogicalProofPdf();
        var layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };
        PdfCore.PdfLogicalDocument logicalDocument = PdfCore.PdfLogicalDocument.Load(pdf, layoutOptions);

        using var semanticWordStream = new MemoryStream();
        var semanticWordOptions = new PdfWordReadOptions();
        PdfWordConversionResult semanticWordResult = logicalDocument.ToWordDocumentResult(semanticWordOptions);
        using (OfficeIMO.Word.WordDocument semanticWordDocument = semanticWordResult.Value) {
            semanticWordDocument.Save(semanticWordStream);
        }

        using var excelStream = new MemoryStream();
        PdfExcelTableImportReport excelReport = PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            logicalDocument,
            excelStream,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false
            });

        using var powerPointStream = new MemoryStream();
        PdfPowerPointTableImportReport powerPointReport = PowerPointPdfConverterExtensions.SaveTablesAsPowerPoint(
            logicalDocument,
            powerPointStream,
            new PdfPowerPointTableImportOptions());

        PdfExcelTableImportEntry excelResult = Assert.Single(excelReport.Entries);
        PdfPowerPointTableImportEntry powerPointResult = Assert.Single(powerPointReport.Entries);

        Assert.Equal(3, excelResult.ColumnCount);
        Assert.Equal(3, powerPointResult.ColumnCount);
        Assert.Equal(2, excelResult.RowCount);
        Assert.Equal(2, powerPointResult.RowCount);
        Assert.True(semanticWordStream.Length > 0);
        Assert.True(excelStream.Length > 0);
        Assert.True(powerPointStream.Length > 0);
        using (WordprocessingDocument semanticWordPackage = WordprocessingDocument.Open(new MemoryStream(semanticWordStream.ToArray()), false)) {
            Assert.Single(semanticWordPackage.MainDocumentPart!.ImageParts);
            Body body = semanticWordPackage.MainDocumentPart.Document.Body!;
            Assert.NotEmpty(body.Descendants<Table>());
            Hyperlink internalLink = Assert.Single(body.Descendants<Hyperlink>(), link => !string.IsNullOrWhiteSpace(link.Anchor?.Value));
            string anchor = Assert.IsType<string>(internalLink.Anchor?.Value);
            Assert.StartsWith("OfficeIMO_Pdf_Dest_Details", anchor, StringComparison.Ordinal);
            Assert.Contains(body.Descendants<BookmarkStart>(), bookmark => bookmark.Name?.Value == anchor);
        }

        Assert.Contains(semanticWordResult.Report.Warnings, warning => warning.Code == "PdfImageEmbedded");
        Assert.DoesNotContain(semanticWordResult.Report.Warnings, warning => warning.Code == "PdfImagePlaceholder");
        Assert.Contains(semanticWordResult.Report.Warnings, warning => warning.Code == "PdfUriLinkReconstructed");
        Assert.Contains(semanticWordResult.Report.Warnings, warning => warning.Code == "PdfInternalLinkReconstructed");
        Assert.DoesNotContain(semanticWordResult.Report.Warnings, warning => warning.Code == "PdfLinkAnnotationNotReconstructed");

        var summary = new {
            scenario = "pdf-table-import-editable-office",
            semanticWordWarningCodes = semanticWordResult.Report.Warnings.Select(warning => warning.Code).Distinct(StringComparer.Ordinal).ToArray(),
            excel = excelResult,
            powerPoint = powerPointResult
        };

        WriteReviewArtifact("pdf-table-import-source.pdf", pdf);
        WriteReviewArtifact("pdf-semantic-import-word.docx", semanticWordStream.ToArray());
        WriteReviewArtifact("pdf-table-import-excel.xlsx", excelStream.ToArray());
        WriteReviewArtifact("pdf-table-import-powerpoint.pptx", powerPointStream.ToArray());
        WriteReviewArtifact("pdf-table-import-editable-office-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static string RequireString(JsonElement element, string propertyName) {
        string? value = element.GetProperty(propertyName).GetString();
        Assert.False(string.IsNullOrWhiteSpace(value), propertyName + " cannot be empty.");
        return value!;
    }

    private static bool HasIssueFeature(JsonElement row, string feature) {
        if (!row.TryGetProperty("Issues", out JsonElement issues) || issues.ValueKind != JsonValueKind.Array) {
            return false;
        }

        return issues.EnumerateArray().Any(issue => RequireString(issue, "Feature") == feature);
    }

    private static IReadOnlyList<string> ReadStringArray(JsonElement element, string propertyName) {
        var values = new List<string>();
        foreach (JsonElement item in element.GetProperty(propertyName).EnumerateArray()) {
            string? value = item.GetString();
            if (!string.IsNullOrWhiteSpace(value)) {
                values.Add(value!);
            }
        }

        return values;
    }

    private static void AssertPremiumQualityContract(JsonElement qualityContract, ISet<string> scenarioIds) {
        Assert.DoesNotContain("TestimoX", RequireString(qualityContract, "goal"), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("OfficeIMO.Pdf", RequireString(qualityContract, "runtimeOwnership"), StringComparison.Ordinal);
        Assert.Contains("thin", RequireString(qualityContract, "runtimeOwnership"), StringComparison.OrdinalIgnoreCase);
        Assert.NotEmpty(ReadStringArray(qualityContract, "nonGoals"));
        Assert.Contains("No TestimoX-specific rendering logic.", ReadStringArray(qualityContract, "nonGoals"));

        IReadOnlyList<string> requiredProof = ReadStringArray(qualityContract, "requiredProof");
        Assert.Contains("artifact-hash", requiredProof);
        Assert.Contains("warning-contract", requiredProof);
        Assert.Contains("accepted-degradation-policy", requiredProof);
        Assert.Contains("engine-owner", requiredProof);
        Assert.Contains("thin-adapter-boundary", requiredProof);

        var fidelityTiers = new HashSet<string>(ReadStringArray(qualityContract, "fidelityTiers"), StringComparer.Ordinal);
        Assert.Contains("premium-core", fidelityTiers);
        Assert.Contains("premium-adapter", fidelityTiers);
        Assert.Contains("accepted-degradation", fidelityTiers);

        var qualityIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonElement quality in qualityContract.GetProperty("scenarioQuality").EnumerateArray()) {
            string id = RequireString(quality, "id");
            Assert.True(qualityIds.Add(id), "Scenario quality ids must be unique. Duplicate: " + id);
            Assert.Contains(id, scenarioIds);
            Assert.Contains(RequireString(quality, "tier"), fidelityTiers);
            Assert.False(string.IsNullOrWhiteSpace(RequireString(quality, "owner")));
            Assert.NotEmpty(ReadStringArray(quality, "closureFocus"));
        }

        Assert.Equal(scenarioIds.OrderBy(id => id, StringComparer.Ordinal), qualityIds.OrderBy(id => id, StringComparer.Ordinal));
        Assert.Contains("pdf-rewrite-preservation-matrix", scenarioIds);
        Assert.Contains("pdf-redaction-removal-proof", scenarioIds);
        Assert.Contains("pdf-form-appearance-semantics", scenarioIds);
        Assert.Contains("pdf-provider-shaped-text", scenarioIds);

        IReadOnlyList<string> knownLimits = ReadStringArray(qualityContract, "knownLimits");
        Assert.Contains(knownLimits, item => item.Contains("IOfficeTextShapingProvider", StringComparison.Ordinal));
        Assert.Contains(knownLimits, item => item.Contains("NativeFontFamilySlotExhausted", StringComparison.Ordinal));
        Assert.Contains(knownLimits, item => item.Contains("OneNote", StringComparison.Ordinal));
        Assert.Contains(knownLimits, item => item.Contains("external validator", StringComparison.OrdinalIgnoreCase));
    }

    private static void AssertConverterCatalog(JsonElement converterCatalog, ISet<string> scenarioIds) {
        string repositoryRoot = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(GetManifestPath())!, ".."));
        string workflow = File.ReadAllText(Path.Combine(repositoryRoot, ".github", "workflows", "pdf-visual-review-gallery.yml"));
        var catalogIds = new HashSet<string>(StringComparer.Ordinal);
        var adapters = new HashSet<string>(StringComparer.Ordinal);

        foreach (JsonElement entry in converterCatalog.EnumerateArray()) {
            string id = RequireString(entry, "id");
            Assert.True(catalogIds.Add(id), "Converter catalog ids must be unique. Duplicate: " + id);
            Assert.True(adapters.Add(RequireString(entry, "adapter")), "Each PDF adapter must have one catalog owner entry.");
            Assert.NotEmpty(ReadStringArray(entry, "sourceFormats"));
            Assert.NotEmpty(ReadStringArray(entry, "extensionTypes"));
            Assert.False(string.IsNullOrWhiteSpace(RequireString(entry, "optionsType")));
            Assert.False(string.IsNullOrWhiteSpace(RequireString(entry, "conversionMode")));

            string projectPath = RequireString(entry, "projectPath");
            Assert.True(File.Exists(Path.Combine(repositoryRoot, projectPath.Replace('/', Path.DirectorySeparatorChar))),
                "Catalog project does not exist: " + projectPath);

            IReadOnlyList<string> ownerPaths = ReadStringArray(entry, "ownerPaths");
            Assert.NotEmpty(ownerPaths);
            foreach (string ownerPath in ownerPaths) {
                Assert.Contains("'" + ownerPath + "'", workflow, StringComparison.Ordinal);
            }

            IReadOnlyList<string> adapterScenarioIds = ReadStringArray(entry, "scenarioIds");
            Assert.NotEmpty(adapterScenarioIds);
            foreach (string scenarioId in adapterScenarioIds) Assert.Contains(scenarioId, scenarioIds);
        }

        Assert.Equal(new[] { "asciidoc", "excel", "html", "latex", "markdown", "onenote", "powerpoint", "rtf", "word" },
            catalogIds.OrderBy(id => id, StringComparer.Ordinal));
    }

    private static void AssertCompositionRoutes(JsonElement compositionRoutes) {
        JsonElement[] routes = compositionRoutes.EnumerateArray().ToArray();
        Assert.NotEmpty(routes);

        string[] expected = {
            "email-document",
            "epub-book",
            "opendocument-presentation-via-powerpoint",
            "opendocument-spreadsheet-via-excel",
            "opendocument-text-via-word",
            "visio-diagram"
        };
        Assert.Equal(expected, routes
            .Select(route => RequireString(route, "id"))
            .OrderBy(id => id, StringComparer.Ordinal)
            .ToArray());

        foreach (JsonElement route in routes) {
            Assert.NotEmpty(ReadStringArray(route, "sourceFormats"));
            Assert.NotEmpty(ReadStringArray(route, "stages"));
            Assert.False(string.IsNullOrWhiteSpace(RequireString(route, "diagnosticContract")));
            string status = RequireString(route, "status");
            Assert.True(
                status == "manual-loss-aware-composition" || status == "not-yet-direct",
                "Unknown composition-route status: " + status);
        }
    }

    private static RtfDocument CreateRtfRoundtripDocument() {
        RtfDocument document = RtfDocument.Create();
        document.Info.Title = "RTF PDF Roundtrip Gate";
        document.Info.Author = "OfficeIMO";
        document.Info.Subject = "RTF PDF semantic proof";
        document.Info.Keywords = "rtf,pdf,roundtrip";
        document.PageSetup.SetPaperSize(7200, 7200);
        document.PageSetup.SetMargins(leftTwips: 720, rightTwips: 720, topTwips: 720, bottomTwips: 720);

        int accentColor = document.AddColor(68, 114, 196);
        document.AddParagraph("RTF PDF Roundtrip Gate").SetStyle(1);

        RtfParagraph rich = document.AddParagraph();
        rich.AddText("Rich run marker ");
        rich.AddText("bold").SetBold().SetForegroundColor(accentColor);
        rich.AddText(" and plain text.");

        document.AddParagraph("Review bullet").SetList(kind: RtfListKind.Bullet).SetIndentation(leftTwips: 720, firstLineTwips: -360);

        RtfTable table = document.AddTable(2, 2);
        table.Rows[0].RepeatHeader = true;
        table.Rows[0].Cells[0].AddParagraph("Area");
        table.Rows[0].Cells[1].AddParagraph("Status");
        table.Rows[1].Cells[0].AddParagraph("Table cell marker");
        table.Rows[1].Cells[1].AddParagraph("Ready");

        RtfParagraph secondPage = document.AddParagraph("Imported second page marker");
        secondPage.PageBreakBefore = true;

        return document;
    }

    private static byte[] CreateLogicalProofPdf() {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                CreateOutlineFromHeadings = true,
                PageWidth = 420,
                PageHeight = 460,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "Logical PDF sample", author: "OfficeIMO")
            .H1("Logical Heading", linkUri: "https://example.com/logical-proof", linkContents: "Logical PDF sample")
            .Paragraph(paragraph => paragraph.Text("Logical readback marker."))
            .Bullets(new[] { "Detected logical bullet" })
            .H2("Jump to details", linkDestinationName: "Details", linkContents: "Jump to details")
            .Bookmark("Details")
            .H2("Details")
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
            .Image(PdfPngTestImages.CreateRgbPng(1, 1), 24, 24, alternativeText: "Logical proof pixel")
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

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreatePdfToHtmlActiveContentProofPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /Fit] /Names << /JavaScript << /Names [(Open) 6 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Contents 4 0 R /Annots [5 0 R] /AA << /O 7 0 R >> >>",
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
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateDirectDestinationLinkPdf() {
        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 5 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Annots [4 0 R] >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Annot /Subtype /Link /Rect [40 160 180 182] /Contents (Jump to page two) /A << /S /GoTo /D [5 0 R /FitR 10 20 90 144] >> >>",
            "endobj",
            "5 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateExcelDashboardReportPdf() {
        string tempDirectory = CreateTemporaryDirectory("OfficeIMOPdfExcelDashboard");
        try {
            string workbookPath = Path.Combine(tempDirectory, "dashboard.xlsx");
            using ExcelDocument document = ExcelDocument.Create(workbookPath, "Dashboard");
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Excel Dashboard PDF Gate");
            sheet.Cell(2, 1, "Pipeline risk");
            sheet.Cell(3, 1, "Channel");
            sheet.Cell(3, 2, "Actual");
            sheet.Cell(3, 3, "Target");
            sheet.Cell(3, 4, "Risk");
            sheet.Cell(4, 1, "Renewals");
            sheet.Cell(4, 2, 128);
            sheet.Cell(4, 3, 120);
            sheet.Cell(4, 4, "Low");
            sheet.Cell(5, 1, "New business");
            sheet.Cell(5, 2, 92);
            sheet.Cell(5, 3, 105);
            sheet.Cell(5, 4, "Medium");
            sheet.Cell(6, 1, "Services");
            sheet.Cell(6, 2, 76);
            sheet.Cell(6, 3, 70);
            sheet.Cell(6, 4, "Low");
            sheet.Cell(7, 1, "Expansion");
            sheet.Cell(7, 2, 54);
            sheet.Cell(7, 3, 65);
            sheet.Cell(7, 4, "High");
            sheet.Cell(9, 1, "Dashboard note");
            sheet.Cell(9, 2, "Charts, image anchors, print area, and conditional formats stay reviewable.");
            sheet.SetColumnWidth(1, 18);
            sheet.SetColumnWidth(2, 12);
            sheet.SetColumnWidth(3, 12);
            sheet.SetColumnWidth(4, 14);
            sheet.SetColumnWidth(5, 16);
            sheet.AddConditionalColorScale("B4:B7", "FFFFF2CC", "FF70AD47");
            sheet.AddConditionalDataBar("C4:C7", "FF5B9BD5");
            sheet.AddImage(10, 1, PdfPngTestImages.CreateRgbPng(2, 2), "image/png", widthPixels: 36, heightPixels: 24, name: "Dashboard badge", altText: "Dashboard badge");
            sheet.AddChartFromRange("A3:C7", row: 1, column: 5, widthPixels: 320, heightPixels: 190, type: ExcelChartType.ColumnClustered, title: "KPI Trend");
            sheet.SetHeaderFooter(headerCenter: "Excel Dashboard PDF Gate", footerRight: "Page &P of &N");
            sheet.SetPageSetup(fitToWidth: 1U, fitToHeight: 1U);
            document.SetPrintArea(sheet, "A1:H14", save: false);
            document.Save();

            return document.ToPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 3,
                PageSize = new PdfCore.PageSize(560, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        } finally {
            Directory.Delete(tempDirectory, recursive: true);
        }
    }

    private static byte[] CreatePowerPointLayoutThemeGroupsPdf() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        presentation.SetThemeColor(PowerPointThemeColor.Accent1, "1D4ED8");
        presentation.SetThemeColor(PowerPointThemeColor.Accent2, "16A34A");
        PowerPointSlide slide = presentation.AddSlide();
        slide.SetBackgroundGradient("172554", "38BDF8", 35D);
        PowerPointTextBox title = slide.AddTextBoxPoints("Layout Theme Group Gate", 18, 8, 220, 34);
        title.FontSize = 14;
        title.Color = "FFFFFF";
        PowerPointTextBox marker = slide.AddTextBoxPoints("Grouped transform marker", 20, 58, 150, 20);
        marker.FontSize = 9;
        marker.Color = "0F172A";
        PowerPointAutoShape preset = slide.AddShapePoints(DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Triangle, 176, 44, 36, 36);
        preset.FillColor = "1F4E79";
        preset.OutlineColor = "1F4E79";
        PowerPointAutoShape first = slide.AddRectanglePoints(20, 20, 30, 20);
        first.FillColor = "FF0000";
        PowerPointAutoShape second = slide.AddRectanglePoints(60, 20, 30, 20);
        second.FillColor = "00AA00";
        slide.GroupShapes(new PowerPointShape[] { first, second }, "Dashboard group");
        DocumentFormat.OpenXml.Presentation.GroupShape group = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!
            .Elements<DocumentFormat.OpenXml.Presentation.GroupShape>()
            .Single();
        TransformGroup transform = group.GroupShapeProperties!.TransformGroup!;
        transform.Extents!.Cx = PowerPointUnits.FromPoints(140);
        transform.Extents.Cy = PowerPointUnits.FromPoints(40);
        transform.ChildExtents!.Cx = PowerPointUnits.FromPoints(70);
        transform.ChildExtents.Cy = PowerPointUnits.FromPoints(20);
        slide.SlidePart.Slide.Save();

        var options = new PowerPointPdfSaveOptions();
        PdfCore.PdfDocumentConversionResult result = presentation.ToPdfDocumentResult(options);
        byte[] pdf = result.ToBytes();
        Assert.False(result.HasWarnings);
        Assert.Equal("1D4ED8", presentation.GetThemeColor(PowerPointThemeColor.Accent1));
        return pdf;
    }

    private static byte[] CreateLogicalDiagnosticsPdf() {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 460,
                PageHeight = 380,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Meta(title: "Revenue Readback Diagnostics", author: "OfficeIMO")
            .H1("Revenue Readback Diagnostics", linkUri: "https://example.com/pdf-logical-diagnostics", linkContents: "Logical diagnostics")
            .Paragraph(paragraph => paragraph.Text("Image geometry and table confidence marker."))
            .Table(new[] {
                new[] { "Metric", "Score", "Owner" },
                new[] { "Renewal quality", "97", "Finance" },
                new[] { "Pipeline coverage", "84", "Sales" },
                new[] { "Risk burn-down", "76", "Operations" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 150, 70, 110 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .Image(PdfPngTestImages.CreateRgbPng(3, 2), 48, 32, alternativeText: "Wide diagnostics badge")
            .Image(PdfPngTestImages.CreateRgbPng(2, 3), 32, 48, alternativeText: "Tall diagnostics badge")
            .ToBytes();
    }

    private static byte[] CreateReaderOcrHandoffCorpusPdf() {
        return PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 300,
                PageHeight = 220,
                MarginLeft = 24,
                MarginRight = 24,
                MarginTop = 24,
                MarginBottom = 24
            })
            .Image(PdfPngTestImages.CreateRgbPng(3, 2), 180, 120, alternativeText: "Scanned statement page")
            .ToBytes();
    }

    private static byte[] CreateReaderXfaFormCorpusPdf() {
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

    private static byte[] CreateCatalogActionsMultiPagePdf() {
        string firstContent = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "50 180 Td",
            "(Catalog action page 1) Tj",
            "ET"
        });
        string secondContent = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "50 180 Td",
            "(Catalog action page 2) Tj",
            "ET"
        });

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /Names << /JavaScript << /Names [(First) 7 0 R (Second) 8 0 R] >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Resources << /Font << /F1 9 0 R >> >> /Contents 5 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Resources << /Font << /F1 9 0 R >> >> /Contents 6 0 R >>",
            "endobj",
            "5 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(firstContent).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            firstContent,
            "endstream",
            "endobj",
            "6 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(secondContent).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            secondContent,
            "endstream",
            "endobj",
            "7 0 obj",
            "<< /S /JavaScript /JS (app.alert('first')) >>",
            "endobj",
            "8 0 obj",
            "<< /S /JavaScript /JS (app.alert('second')) >>",
            "endobj",
            "9 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 10 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateReaderDegradationCorpusPdf() {
        string content = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "50 180 Td",
            "(Reader Degradation Corpus) Tj",
            "0 -18 Td",
            "(Accepted degradation marker) Tj",
            "0 -18 Td",
            "(Form and active-content marker) Tj",
            "ET"
        });

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [6 0 R] /DA (/Helv 10 Tf 0 g) /DR << /Font << /Helv 7 0 R >> >> >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Resources << /Font << /F1 7 0 R >> >> /Contents 4 0 R /Annots [5 0 R 6 0 R] /AA << /O 8 0 R >> >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Link /Rect [50 132 190 148] /Contents (Review link) /A << /S /URI /URI (https://example.com/reader-degradation) >> >>",
            "endobj",
            "6 0 obj",
            "<< /Type /Annot /Subtype /Widget /FT /Tx /T (Corpus.Contact) /V (review@example.com) /Rect [50 82 220 104] /P 3 0 R >>",
            "endobj",
            "7 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "8 0 obj",
            "<< /S /JavaScript /JS (app.alert('OfficeIMO')) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 9 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateReaderHostileActionCorpusPdf() {
        string content = string.Join("\n", new[] {
            "BT",
            "/F1 12 Tf",
            "50 180 Td",
            "(Hostile Action Corpus) Tj",
            "0 -18 Td",
            "(Passive diagnostics marker) Tj",
            "ET"
        });

        string pdf = string.Join("\n", new[] {
            "%PDF-1.7",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R /OpenAction [3 0 R /Fit] /Names << /JavaScript << /Kids [13 0 R 14 0 R] >> >> /AA << /WC 15 0 R >> >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 320 220] /Resources << /Font << /F1 12 0 R >> >> /Contents 4 0 R /Annots [5 0 R 10 0 R] /AA << /O 7 0 R /C 8 0 R >> >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Annot /Subtype /Link /Rect [50 132 210 150] /Contents (Submit review) /A 9 0 R >>",
            "endobj",
            "6 0 obj",
            "<< /S /JavaScript /JS (app.alert('catalog')) >>",
            "endobj",
            "7 0 obj",
            "<< /S /JavaScript /JS (app.alert('page-open')) >>",
            "endobj",
            "8 0 obj",
            "<< /S /Launch /F (tool.exe) >>",
            "endobj",
            "9 0 obj",
            "<< /S /SubmitForm /F (https://example.com/submit) /Next 11 0 R >>",
            "endobj",
            "10 0 obj",
            "<< /Type /Annot /Subtype /Text /Rect [50 96 80 126] /Contents (Launch review note) /AA << /E << /S /Launch /F (hover.exe) >> >> >>",
            "endobj",
            "11 0 obj",
            "<< /S /ImportData /F (payload.fdf) >>",
            "endobj",
            "12 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "13 0 obj",
            "<< /Names [(Startup) 6 0 R] >>",
            "endobj",
            "14 0 obj",
            "<< /Names [(Deferred) 16 0 R] >>",
            "endobj",
            "15 0 obj",
            "<< /S /JavaScript /JS (app.alert('catalog-close')) /Next [17 0 R 18 0 R] >>",
            "endobj",
            "16 0 obj",
            "<< /S /JavaScript /JS (app.alert('deferred')) >>",
            "endobj",
            "17 0 obj",
            "<< /S /ImportData /F (name-tree-data.fdf) >>",
            "endobj",
            "18 0 obj",
            "<< /S /Movie /T (preview.mov) >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 19 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateReaderHostileLayoutCorpusPdf() {
        string content = string.Join("\n", new[] {
            "BT",
            "/F1 14 Tf",
            "50 720 Td",
            "(Hostile Layout Corpus) Tj",
            "ET",
            "BT",
            "/F1 10 Tf",
            "50 680 Td",
            "(Left column marker) Tj",
            "0 -14 Td",
            "(Left column value 42) Tj",
            "ET",
            "BT",
            "/F1 10 Tf",
            "185 680 Td",
            "(Right column marker) Tj",
            "0 -14 Td",
            "(Right column value 84) Tj",
            "ET",
            "BT",
            "/F1 10 Tf",
            "0 1 -1 0 330 610 Tm",
            "(Rotated note marker) Tj",
            "ET",
            "q",
            "36 12 18 24 260 84 cm",
            "/Im1 Do",
            "Q"
        });

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 420 760] /Resources << /Font << /F1 5 0 R >> /XObject << /Im1 6 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "6 0 obj",
            "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Length 3 >>",
            "stream",
            "abc",
            "endstream",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 7 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
    }

    private static byte[] CreateReaderHostileTableCorpusPdf() {
        string content = string.Join("\n", new[] {
            "BT",
            "/F1 14 Tf",
            "50 720 Td",
            "(Hostile Table Corpus) Tj",
            "ET",
            "BT",
            "/F1 10 Tf",
            "50 690 Td",
            "(Jittered table marker) Tj",
            "ET",
            "BT",
            "/F1 10 Tf",
            "56 646 Td",
            "(Alpha team) Tj",
            "132 0 Td",
            "(Ops) Tj",
            "74 0 Td",
            "(72) Tj",
            "ET",
            "BT",
            "/F1 10 Tf",
            "58 630 Td",
            "(Beta team) Tj",
            "125 0 Td",
            "(Sales) Tj",
            "79 0 Td",
            "(85) Tj",
            "ET",
            "BT",
            "/F1 10 Tf",
            "54 614 Td",
            "(Gamma team) Tj",
            "138 0 Td",
            "(Risk) Tj",
            "69 0 Td",
            "(91) Tj",
            "ET"
        });

        string pdf = string.Join("\n", new[] {
            "%PDF-1.4",
            "1 0 obj",
            "<< /Type /Catalog /Pages 2 0 R >>",
            "endobj",
            "2 0 obj",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "endobj",
            "3 0 obj",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 420 760] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
            "endobj",
            "4 0 obj",
            "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(System.Globalization.CultureInfo.InvariantCulture) + " >>",
            "stream",
            content,
            "endstream",
            "endobj",
            "5 0 obj",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
            "endobj",
            "trailer",
            "<< /Root 1 0 R /Size 6 >>",
            "%%EOF"
        }) + "\n";

        return Encoding.ASCII.GetBytes(pdf);
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

    private static string CreateCssResourcePolicyHtml(string stylesheetUri) {
        string pixel = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(1, 1));
        return $$"""
<html>
<head>
  <link rel="stylesheet" href="{{stylesheetUri}}">
  <link rel="stylesheet" href="https://blocked.example.test/policy.css">
</head>
<body>
  <h1>HTML CSS Resource Policy Gate</h1>
  <p class="policy-note">Local stylesheet marker</p>
  <p><a href="https://example.com/pdf-resource-policy">Policy link</a></p>
  <p>Blocked remote stylesheet marker</p>
  <p><img src="data:image/png;base64,{{pixel}}" alt="Policy pixel" width="24" height="24"></p>
  <table>
    <tr><th>Resource</th><th>Policy</th></tr>
    <tr><td>file stylesheet</td><td>allowed</td></tr>
    <tr><td>blocked remote stylesheet</td><td>diagnostic</td></tr>
  </table>
</body>
</html>
""";
    }

    private static string CreateInvoiceStatementMarkdown() {
        return """
---
title: OfficeIMO invoice statement proof
author: OfficeIMO
tags: [pdf, invoice, statement]
pdfTheme: report
---

# Invoice Statement INV-2026-0042

Bill to: Contoso Finance Review

| Service | Period | Quantity | Amount |
| --- | --- | ---: | ---: |
| Managed PDF conversion review | 2026-Q2 | 1 | 1200.00 |
| Table extraction proof pack | 2026-Q2 | 2 | 450.00 |
| Visual gallery artifacts | 2026-Q2 | 1 | 175.00 |

| Summary | Amount |
| --- | ---: |
| Subtotal | 2275.00 |
| Tax | 523.25 |
| Amount due | 2798.25 |

- Payment terms: Net 14
- Remittance reference: INV-2026-0042
- Review note: totals and right-aligned numeric columns must remain inspectable.

Thank you for reviewing the OfficeIMO PDF conversion statement.
""";
    }

    private static IReadOnlyList<DrawingCore.OfficeShapedGlyph> CreateTrueTypeGlyphMap(string text, PdfCore.PdfTrueTypeFontProgram fontProgram) {
        var glyphs = new List<DrawingCore.OfficeShapedGlyph>();
        for (int index = 0; index < text.Length;) {
            int scalarStart = index;
            int scalar = ReadScalar(text, ref index);
            if (!fontProgram.TryGetGlyphId(scalar, out int glyphId)) {
                throw new InvalidOperationException("The selected TrueType font does not cover " + scalar.ToString("X", System.Globalization.CultureInfo.InvariantCulture) + ".");
            }

            glyphs.Add(new DrawingCore.OfficeShapedGlyph(glyphId, char.ConvertFromUtf32(scalar), scalarStart));
        }

        return glyphs;
    }

    private static bool TryCreateCffOfficeLigatureGlyphs(PdfCore.PdfOpenTypeCffFontProgram fontProgram, out IReadOnlyList<DrawingCore.OfficeShapedGlyph>? glyphs) {
        glyphs = null;
        if (!fontProgram.TryGetGlyphId('o', out int oGlyphId) ||
            !fontProgram.TryGetGlyphId(0xFB03, out int ffiGlyphId) ||
            !fontProgram.TryGetGlyphId('c', out int cGlyphId) ||
            !fontProgram.TryGetGlyphId('e', out int eGlyphId)) {
            return false;
        }

        glyphs = new[] {
            new DrawingCore.OfficeShapedGlyph(oGlyphId, "o", 0),
            new DrawingCore.OfficeShapedGlyph(ffiGlyphId, "ffi", 1),
            new DrawingCore.OfficeShapedGlyph(cGlyphId, "c", 4),
            new DrawingCore.OfficeShapedGlyph(eGlyphId, "e", 5)
        };
        return true;
    }

    private static int ReadScalar(string text, ref int index) {
        char ch = text[index++];
        if (char.IsHighSurrogate(ch) && index < text.Length && char.IsLowSurrogate(text[index])) {
            return char.ConvertToUtf32(ch, text[index++]);
        }

        return ch;
    }

    private sealed class ManifestTextShapingProvider : DrawingCore.IOfficeTextShapingProvider {
        private readonly string _trueTypeText;
        private readonly IReadOnlyList<DrawingCore.OfficeShapedGlyph> _trueTypeGlyphs;
        private readonly string _cffText;
        private readonly IReadOnlyList<DrawingCore.OfficeShapedGlyph> _cffGlyphs;

        public ManifestTextShapingProvider(
            string trueTypeText,
            IReadOnlyList<DrawingCore.OfficeShapedGlyph> trueTypeGlyphs,
            string cffText,
            IReadOnlyList<DrawingCore.OfficeShapedGlyph> cffGlyphs) {
            _trueTypeText = trueTypeText;
            _trueTypeGlyphs = trueTypeGlyphs;
            _cffText = cffText;
            _cffGlyphs = cffGlyphs;
        }

        public int CallCount { get; private set; }
        public int TrueTypeCalls { get; private set; }
        public int OpenTypeCffCalls { get; private set; }

        public DrawingCore.OfficeTextShapingResult? ShapeText(DrawingCore.OfficeTextShapingRequest request) {
            CallCount++;
            if (!request.IsOpenTypeCff && string.Equals(request.Text, _trueTypeText, StringComparison.Ordinal)) {
                TrueTypeCalls++;
                return new DrawingCore.OfficeTextShapingResult(_trueTypeGlyphs);
            }

            if (request.IsOpenTypeCff && string.Equals(request.Text, _cffText, StringComparison.Ordinal)) {
                OpenTypeCffCalls++;
                return new DrawingCore.OfficeTextShapingResult(_cffGlyphs);
            }

            return null;
        }
    }

    private static string CreateTemporaryDirectory(string prefix) {
        string path = Path.Combine(Path.GetTempPath(), prefix + "-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }

    private static void WriteReviewArtifact(string fileName, byte[] bytes) {
        string? outputDirectory = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT");
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            return;
        }

        Directory.CreateDirectory(outputDirectory);
        File.WriteAllBytes(Path.Combine(outputDirectory, fileName), bytes);
    }

    private static void AssertReviewArtifactPrerequisite(string message) {
        if (!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_REVIEW_OUTPUT"))) {
            Assert.Fail(message);
        }
    }

    private static string GetManifestPath() {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory != null) {
            string candidate = Path.Combine(directory.FullName, "Docs", "pdf-conversion-scenarios.json");
            if (File.Exists(candidate)) {
                return candidate;
            }

            directory = directory.Parent;
        }

        throw new FileNotFoundException("Could not locate Docs/pdf-conversion-scenarios.json from test runtime base directory.");
    }

}
