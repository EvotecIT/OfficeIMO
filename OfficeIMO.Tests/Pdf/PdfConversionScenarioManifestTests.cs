using System.Text;
using System.Text.Json;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using TransformGroup = DocumentFormat.OpenXml.Drawing.TransformGroup;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConversionScenarioManifestTests {
    private static readonly string[] RequiredPaths = {
        "word-to-pdf",
        "excel-to-pdf",
        "excel-dashboard-to-pdf",
        "markdown-to-pdf",
        "markdown-invoice-to-pdf",
        "html-to-pdf",
        "html-css-resource-policy-to-pdf",
        "html-pdf-roundtrip",
        "powerpoint-to-pdf",
        "powerpoint-layout-theme-groups-to-pdf",
        "pdf-to-logical",
        "pdf-logical-diagnostics-readback",
        "pdf-reader-degradation-corpus",
        "pdf-reader-hostile-layout-corpus",
        "pdf-reader-hostile-table-corpus",
        "pdf-to-html",
        "pdf-to-editable-office-tables"
    };

    [Fact]
    public void PdfConversionManifest_CoversEverySupportedConversionPathWithObservableProof() {
        using JsonDocument document = JsonDocument.Parse(File.ReadAllText(GetManifestPath()));
        JsonElement root = document.RootElement;
        Assert.Equal(1, root.GetProperty("version").GetInt32());

        JsonElement scenarios = root.GetProperty("scenarios");
        Assert.True(scenarios.GetArrayLength() >= RequiredPaths.Length);

        var seenPaths = new HashSet<string>(StringComparer.Ordinal);
        var seenIds = new HashSet<string>(StringComparer.Ordinal);
        foreach (JsonElement scenario in scenarios.EnumerateArray()) {
            string id = RequireString(scenario, "id");
            Assert.True(seenIds.Add(id), "Scenario ids must be unique. Duplicate: " + id);

            string path = RequireString(scenario, "path");
            seenPaths.Add(path);
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

        foreach (string requiredPath in RequiredPaths) {
            Assert.Contains(requiredPath, seenPaths);
        }
    }

    [Fact]
    public void HtmlDocumentProfile_ProducesManifestedReviewProof() {
        const string linkUri = "https://example.com/pdf-conversion-manifest";
        byte[] pdf = CreatePracticalHtmlSample(linkUri).SaveAsPdf(HtmlPdfSaveOptions.CreateDocumentProfile());
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

        string html = PdfHtmlConverter.ToHtml(logical, options);

        Assert.Equal(1, logical.PageCount);
        Assert.Contains(logical.Headings, heading => heading.Text == "Logical Heading");
        Assert.Contains(logical.ListItems, item => item.Text == "Detected logical bullet");
        Assert.NotEmpty(logical.Tables);
        Assert.Contains(logical.GetLinksByUri("https://example.com/logical-proof"), link => link.Contents == "Logical PDF sample");
        Assert.Contains(logical.Images, image => image.Width > 0D && image.Height > 0D);
        Assert.Contains("class=\"pdf-page\" data-page-number=\"1\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-text pdf-heading\"", html, StringComparison.Ordinal);
        Assert.Contains("Logical Heading", html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-image-placeholder\"", html, StringComparison.Ordinal);
        Assert.False(options.ConversionReport.HasWarnings);

        WriteReviewArtifact("pdf-to-html-logical-source.pdf", pdf);
        WriteReviewArtifact("pdf-to-html-positioned-review.html", System.Text.Encoding.UTF8.GetBytes(html));
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

        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

        Assert.Contains("Q2 multilingual revenue report", text, StringComparison.Ordinal);
        Assert.Contains("Zażółć gęślą jaźń", text, StringComparison.Ordinal);
        Assert.Contains("Ελλάδα", text, StringComparison.Ordinal);
        Assert.Contains("Київ", text, StringComparison.Ordinal);

        WriteReviewArtifact("multilingual-business-report.pdf", pdf);
    }

    [Fact]
    public void MarkdownInvoiceStatement_ProducesManifestedReviewProof() {
        byte[] pdf = CreateInvoiceStatementMarkdown().SaveAsPdf(new MarkdownPdfSaveOptions {
            ApplyWordLikeTheme = true,
            Title = "OfficeIMO invoice statement proof",
            Subject = "Invoice and statement conversion proof"
        });

        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

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
    public void ExcelDashboardReport_ProducesManifestedReviewProof() {
        byte[] pdf = CreateExcelDashboardReportPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();

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
        string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
        string raw = Encoding.ASCII.GetString(pdf);

        Assert.True(pdf.Length > 0);
        Assert.Contains("Layout Theme Group Gate", text, StringComparison.Ordinal);
        Assert.Contains("Grouped transform marker", text, StringComparison.Ordinal);
        Assert.Contains("20 100 60 40 re", raw, StringComparison.Ordinal);
        Assert.Contains("100 100 60 40 re", raw, StringComparison.Ordinal);

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

        string html = PdfHtmlConverter.ToHtml(logical, options);
        PdfCore.PdfLogicalTableExtraction extraction = Assert.Single(PdfCore.PdfLogicalTableAnalysis.ExtractTables(logical));
        PdfCore.PdfLogicalTableColumnProfile scoreProfile = Assert.Single(extraction.Data.ColumnProfiles, profile => profile.Name == "Score");
        PdfCore.PdfLogicalImage wideImage = Assert.Single(logical.Images, image => image.Width == 3 && image.Height == 2);
        PdfCore.PdfLogicalImage tallImage = Assert.Single(logical.Images, image => image.Width == 2 && image.Height == 3);
        ReaderChunk readerChunk = Assert.Single(DocumentReaderPdfExtensions.ReadPdf(
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
        Assert.Contains("class=\"pdf-page\" data-page-number=\"1\"", html, StringComparison.Ordinal);
        Assert.Contains("Revenue Readback Diagnostics", html, StringComparison.Ordinal);
        Assert.False(options.ConversionReport.HasWarnings);
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
    public void PdfReaderDegradationCorpus_ProducesManifestedReaderProof() {
        byte[] pdf = CreateReaderDegradationCorpusPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });

        ReaderChunk chunk = Assert.Single(DocumentReaderPdfExtensions.ReadPdf(
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
                chunk.Diagnostics.SelectedAnnotationActionCount
            },
            actions = chunk.Actions!.Select(item => new {
                Scope = item.Scope.ToString(),
                item.ActionType,
                item.Source,
                item.TriggerName,
                item.ActionPath,
                item.PageNumber,
                item.IsChainedAction
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
    public void PdfReaderHostileLayoutCorpus_ProducesManifestedReaderProof() {
        byte[] pdf = CreateReaderHostileLayoutCorpusPdf();
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = false
        });

        ReaderChunk chunk = Assert.Single(DocumentReaderPdfExtensions.ReadPdf(
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

        ReaderChunk chunk = Assert.Single(DocumentReaderPdfExtensions.ReadPdf(
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
        Assert.NotNull(chunk.Tables);
        ReaderTable table = Assert.Single(chunk.Tables!);
        Assert.Equal(new[] { "Column 1", "Column 2", "Column 3" }, table.Columns);
        Assert.NotNull(table.Diagnostics);
        Assert.True(table.Diagnostics!.Confidence >= 0.80D);
        Assert.True(table.Diagnostics.Confidence < 0.95D);
        Assert.Equal(ReaderTableColumnKind.Numeric, table.ColumnProfiles[2].Kind);

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
                chunk.Diagnostics.AverageTableConfidence
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
    public void HtmlCssResourcePolicy_ProducesManifestedReviewProof() {
        string tempDirectory = CreateTemporaryDirectory("OfficeIMOPdfHtmlPolicy");
        try {
            string stylesheetPath = Path.Combine(tempDirectory, "policy.css");
            File.WriteAllText(stylesheetPath, "p.policy-note { color:#123456; }", Encoding.UTF8);
            var options = HtmlPdfSaveOptions.CreateTrustedDocumentProfile();
            options.WordHtmlOptions!.AllowedStylesheetHosts.Add("allowed.example.test");
            options.WordHtmlOptions.MaxCssBytes = 8192;
            options.WordHtmlOptions.MaxTotalCssBytes = 16384;

            byte[] pdf = CreateCssResourcePolicyHtml(new Uri(stylesheetPath).AbsoluteUri).SaveAsPdf(options);
            string text = PdfCore.PdfReadDocument.Load(pdf).ExtractText();
            HtmlPdfResourcePolicySummary policy = options.GetResourcePolicySummary();

            Assert.True(pdf.Length > 0);
            Assert.True(options.WordHtmlOptions.AllowDocumentStylesheetLinks);
            Assert.Contains(Uri.UriSchemeFile, options.WordHtmlOptions.AllowedStylesheetUriSchemes);
            Assert.Contains(options.WordHtmlOptions.Diagnostics, diagnostic => diagnostic.Code == "StylesheetResourceRejectedByPolicy");
            Assert.True(policy.UsesWordHtmlPolicy);
            Assert.True(policy.AllowDocumentStylesheetLinks);
            Assert.Contains(Uri.UriSchemeFile, policy.AllowedStylesheetUriSchemes);
            Assert.Contains("allowed.example.test", policy.AllowedStylesheetHosts);
            Assert.Equal(8192, policy.MaxCssBytes);
            Assert.Equal(16384, policy.MaxTotalCssBytes);
            PdfCore.PdfConversionWarning stylesheetWarning = Assert.Single(options.ConversionReport.Warnings, warning => warning.Code == "StylesheetResourceRejectedByPolicy");
            Assert.Equal("OfficeIMO.Word.Html", stylesheetWarning.Converter);
            Assert.Contains("HTML CSS Resource Policy Gate", text, StringComparison.Ordinal);
            Assert.Contains("Local stylesheet marker", text, StringComparison.Ordinal);
            Assert.Contains("Blocked remote stylesheet marker", text, StringComparison.Ordinal);
            Assert.Contains(PdfCore.PdfImageExtractor.ExtractImages(pdf), image => image.IsImageFile && image.MimeType == "image/png");

            var summary = new {
                scenario = "html-css-resource-policy",
                profile = options.Profile.ToString(),
                policy,
                diagnostics = options.ConversionReport.Warnings.Where(warning => warning.Converter == "OfficeIMO.Word.Html").Select(warning => new {
                    warning.Converter,
                    warning.Code,
                    warning.Source,
                    warning.Message,
                    Severity = warning.Severity.ToString(),
                    warning.Details
                }).ToArray()
            };

            WriteReviewArtifact("html-css-resource-policy.pdf", pdf);
            WriteReviewArtifact("html-css-resource-policy-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions { WriteIndented = true }));
        } finally {
            Directory.Delete(tempDirectory, recursive: true);
        }
    }

    [Fact]
    public void HtmlPdfRoundTripProfiles_ProduceManifestedReviewProof() {
        const string linkUri = "https://example.com/html-pdf-roundtrip";
        HtmlPdfSaveOptions htmlOptions = HtmlPdfSaveOptions.CreateDocumentProfile();
        byte[] pdf = CreatePracticalHtmlSample(linkUri).SaveAsPdf(htmlOptions);
        PdfCore.PdfLogicalDocument logical = PdfCore.PdfLogicalDocument.Load(pdf, new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        });
        var semanticOptions = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.Semantic,
            IncludeLinkAnnotations = true,
            LayoutOptions = new PdfCore.PdfTextLayoutOptions {
                ForceSingleColumn = true
            }
        };
        var positionedOptions = new PdfHtmlSaveOptions {
            Profile = PdfHtmlProfile.PositionedReview,
            IncludeLinkAnnotations = true,
            LayoutOptions = new PdfCore.PdfTextLayoutOptions {
                ForceSingleColumn = true
            }
        };

        PdfHtmlConversionResult semantic = PdfHtmlConverter.ToHtmlResult(logical, semanticOptions);
        PdfHtmlConversionResult positioned = PdfHtmlConverter.ToHtmlResult(logical, positionedOptions);

        Assert.True(pdf.Length > 0);
        Assert.True(logical.PageCount >= 2);
        Assert.Contains(logical.TextBlocks, block => block.Text.Contains("Practical HTML", StringComparison.Ordinal));
        Assert.Contains(logical.GetLinksByUri(linkUri), link => link.Contents == "Report link");
        Assert.Contains(logical.Images, image => image.PlacementCount > 0);
        Assert.Contains("Practical HTML", semantic.Html, StringComparison.Ordinal);
        Assert.Contains("Report link", semantic.Html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-page\" data-page-number=\"1\"", positioned.Html, StringComparison.Ordinal);
        Assert.Contains("class=\"pdf-image-placeholder\"", positioned.Html, StringComparison.Ordinal);
        Assert.Equal(PdfHtmlProfile.Semantic, semantic.Summary.Profile);
        Assert.Equal(PdfHtmlProfile.PositionedReview, positioned.Summary.Profile);
        Assert.True(semantic.Summary.RenderedPageCount >= 2);
        Assert.True(positioned.Summary.RenderedPageCount >= 2);
        Assert.True(positioned.Summary.TextBlockCount > 0);
        Assert.True(positioned.Summary.ImagePlacementCount > 0);
        Assert.True(positioned.Summary.LinkCount > 0);
        Assert.False(semantic.ConversionReport.HasWarnings);
        Assert.False(positioned.ConversionReport.HasWarnings);

        var summary = new {
            scenario = "html-pdf-roundtrip-profile-contract",
            htmlToPdfProfile = HtmlPdfProfileContracts.Get(HtmlPdfProfile.Document),
            pdfToSemanticProfile = PdfHtmlProfileContracts.Get(PdfHtmlProfile.Semantic),
            pdfToPositionedProfile = PdfHtmlProfileContracts.Get(PdfHtmlProfile.PositionedReview),
            htmlToPdfWarnings = htmlOptions.ConversionReport.Warnings.Select(warning => new {
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
        WriteReviewArtifact("html-pdf-roundtrip-semantic.html", Encoding.UTF8.GetBytes(semantic.Html));
        WriteReviewArtifact("html-pdf-roundtrip-positioned.html", Encoding.UTF8.GetBytes(positioned.Html));
        WriteReviewArtifact("html-pdf-roundtrip-summary.json", JsonSerializer.SerializeToUtf8Bytes(summary, new JsonSerializerOptions {
            WriteIndented = true
        }));
    }

    [Fact]
    public void PdfTableImportProfiles_ProduceManifestedEditableOfficeProof() {
        byte[] pdf = CreateLogicalProofPdf();
        var layoutOptions = new PdfCore.PdfTextLayoutOptions {
            ForceSingleColumn = true
        };

        using var wordStream = new MemoryStream();
        IReadOnlyList<PdfWordTableImportResult> wordResults = PdfWordTableConverterExtensions.SavePdfTablesAsWord(
            pdf,
            wordStream,
            new PdfWordTableImportOptions {
                LayoutOptions = layoutOptions
            });

        using var excelStream = new MemoryStream();
        IReadOnlyList<PdfExcelTableImportResult> excelResults = PdfExcelTableConverterExtensions.SavePdfTablesAsExcel(
            pdf,
            excelStream,
            new PdfExcelTableImportOptions {
                LayoutOptions = layoutOptions,
                AutoFitColumns = false
            });

        using var powerPointStream = new MemoryStream();
        IReadOnlyList<PdfPowerPointTableImportResult> powerPointResults = PowerPointPdfConverterExtensions.SavePdfTablesAsPowerPoint(
            pdf,
            powerPointStream,
            new PdfPowerPointTableImportOptions {
                LayoutOptions = layoutOptions
            });

        PdfWordTableImportResult wordResult = Assert.Single(wordResults);
        PdfExcelTableImportResult excelResult = Assert.Single(excelResults);
        PdfPowerPointTableImportResult powerPointResult = Assert.Single(powerPointResults);

        Assert.Equal(3, wordResult.ColumnCount);
        Assert.Equal(3, excelResult.ColumnCount);
        Assert.Equal(3, powerPointResult.ColumnCount);
        Assert.Equal(2, wordResult.RowCount);
        Assert.Equal(2, excelResult.RowCount);
        Assert.Equal(2, powerPointResult.RowCount);
        Assert.True(wordStream.Length > 0);
        Assert.True(excelStream.Length > 0);
        Assert.True(powerPointStream.Length > 0);

        WriteReviewArtifact("pdf-table-import-source.pdf", pdf);
        WriteReviewArtifact("pdf-table-import-word.docx", wordStream.ToArray());
        WriteReviewArtifact("pdf-table-import-excel.xlsx", excelStream.ToArray());
        WriteReviewArtifact("pdf-table-import-powerpoint.pptx", powerPointStream.ToArray());
    }

    private static string RequireString(JsonElement element, string propertyName) {
        string? value = element.GetProperty(propertyName).GetString();
        Assert.False(string.IsNullOrWhiteSpace(value), propertyName + " cannot be empty.");
        return value!;
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

    private static byte[] CreateLogicalProofPdf() {
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
            .H1("Logical Heading", linkUri: "https://example.com/logical-proof", linkContents: "Logical PDF sample")
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
            .Image(PdfPngTestImages.CreateRgbPng(1, 1), 24, 24, alternativeText: "Logical proof pixel")
            .ToBytes();
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
            document.Save(false);

            return document.SaveAsPdf(new ExcelPdfSaveOptions {
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
        PowerPointSlide slide = presentation.Slides[0];
        slide.SetBackgroundGradient("172554", "38BDF8", 35D);
        PowerPointTextBox title = slide.AddTextBoxPoints("Layout Theme Group Gate", 18, 10, 190, 24);
        title.FontSize = 14;
        title.Color = "FFFFFF";
        PowerPointTextBox marker = slide.AddTextBoxPoints("Grouped transform marker", 20, 58, 150, 20);
        marker.FontSize = 9;
        marker.Color = "0F172A";
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
        byte[] pdf = presentation.SaveAsPdf(options);
        Assert.Empty(options.Warnings);
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
