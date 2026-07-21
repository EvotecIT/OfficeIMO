using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfEngineRoadmapFoundationTests {
    [Fact]
    public void BusinessRecipeComponents_ProduceSearchableArtifactsThroughCanonicalFlow() {
        var report = new PdfReportComponent(
            "Operations report",
            "Stable shared-engine summary",
            new[] { new KeyValuePair<string, string?>("Status", "Healthy") },
            new[] { new PdfReportSection("Signals", bullets: new[] { "Consistent", "Evidence-backed" }) });
        var invoice = new PdfInvoiceComponent(
            "INV-42",
            new DateTime(2026, 7, 21),
            new PdfInvoiceParty("Seller Ltd"),
            new PdfInvoiceParty("Customer Ltd"),
            new[] { new PdfInvoiceLine("Engine work", 2M, 50M, 0.20M) },
            "EUR");
        var labels = new PdfLabelSheetComponent(new[] {
            new PdfLabel("Premium", "Batch A", "LBL-001"),
            new PdfLabel("Premium", "Batch B", "LBL-002")
        }, columns: 2);
        var ticket = new PdfTicketComponent("Engine summit", "TICKET-7", venue: "Main hall", holder: "Ada");

        byte[] bytes = PdfDocument.Create()
            .Component(report)
            .Component(invoice)
            .Component(labels)
            .Component(ticket)
            .ToBytes();

        string text = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("Operations report", text, StringComparison.Ordinal);
        Assert.Contains("INV-42", text, StringComparison.Ordinal);
        Assert.Contains("120.00 EUR", text, StringComparison.Ordinal);
        Assert.Contains("LBL-002", text, StringComparison.Ordinal);
        Assert.Contains("TICKET-7", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ContextComponent_ReplaysThroughExistingDeferredPaginationPath() {
        var component = new PageAwareComponent();
        PdfDocument document = PdfDocument.Create(new PdfOptions {
            PageWidth = 200,
            PageHeight = 200,
            MarginLeft = 20,
            MarginTop = 20,
            MarginRight = 20,
            MarginBottom = 20,
            CompressContentStreams = false
        });

        byte[] bytes = document
            .Paragraph(paragraph => paragraph.Text("Page one"))
            .Spacer(100)
            .Component(component, new PdfFlowOptions { MinimumRemainingHeight = 100 })
            .ToBytes();

        Assert.Equal(2, component.Invocations);
        Assert.Contains("Context page 2", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void MutationPortfolio_UsesOnePreflightAcrossInteractivePlans() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Mutation portfolio"))
            .TextField("customer", value: "Ada")
            .ToBytes();
        PdfDocument opened = PdfDocument.Open(bytes);

        PdfMutationPortfolioReport report = opened.AssessMutations(new[] {
            PdfMutationOperation.ModifyAnnotations,
            PdfMutationOperation.ModifyCatalog,
            PdfMutationOperation.FillFormFields
        });

        Assert.Equal(3, report.Plans.Count);
        Assert.All(report.Plans, plan => Assert.Same(report.Preflight, plan.Preflight));
        Assert.Contains(
            report.Get(PdfMutationOperation.ModifyAnnotations).CapabilityRecords,
            record => record.Kind == PdfMutationCapabilityKind.AnnotationChanges);
        Assert.Contains(report.Get(PdfMutationOperation.FillFormFields).CapabilityRecords, record => record.Kind == PdfMutationCapabilityKind.FormChanges);
    }

    [Fact]
    public void MutationPortfolio_SnapshotsOneShotFieldNamesForEveryPlan() {
        byte[] bytes = PdfDocument.Create()
            .TextField("customer", value: "Ada")
            .ToBytes();
        int enumerationCount = 0;

        IEnumerable<string> FieldNames() {
            enumerationCount++;
            yield return "customer";
        }

        PdfMutationPortfolioReport report = PdfDocument.Open(bytes).AssessMutations(
            new[] { PdfMutationOperation.ModifyCatalog, PdfMutationOperation.FillFormFields },
            FieldNames());

        Assert.Equal(2, report.Plans.Count);
        Assert.Equal(1, enumerationCount);
    }

    [Fact]
    public void RenderCompatibility_AssessesSameManifestUsedByPageExport() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Producer compatibility"))
            .ToBytes();

        PdfRenderCompatibilityReport report = PdfDocument.Open(bytes).AssessRenderCompatibility();

        Assert.Same(PdfRenderCapabilities.Current, report.Manifest);
        Assert.Single(report.Pages);
        Assert.Contains("render.resource.font-substitution", report.CapabilityCodes);
        Assert.True(report.HasSimplifications);
        Assert.False(report.IsExactForManagedRenderer);
    }

    [Fact]
    public void Save_ReportsBoundedPayloadSpillWithoutClaimingForwardOnlyLayout() {
        var options = new PdfOptions {
            ObjectBufferMemoryLimitBytes = 0,
            PageContentMemoryLimitBytes = 0
        };
        using var output = new MemoryStream();

        PdfSaveResult result = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.Text("Bounded serialization evidence"))
            .Save(output);

        PdfSerializationReport report = Assert.IsType<PdfSerializationReport>(result.Serialization);
        Assert.True(report.PageContentSpilled);
        Assert.True(report.ObjectBufferSpilled);
        Assert.Equal(0, report.PeakRetainedPageContentBytes);
        Assert.Equal(0, report.PeakRetainedObjectBytes);
        Assert.False(report.FinalArtifactBuffered);
        Assert.False(report.IsForwardOnlyLayout);
        Assert.Equal(result.BytesWritten, report.BytesWritten);
    }

    [Fact]
    public void ReaderProjection_MergesSourceEvidenceAndExplicitEmailAssetPolicy() {
        var source = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Email,
            Source = new OfficeDocumentSource { Title = "Project update", Author = "Ada" },
            Blocks = new[] {
                new OfficeDocumentBlock { Id = "body", Kind = "paragraph", Text = "The release is ready." }
            },
            Assets = new[] {
                new OfficeDocumentAsset { Id = "attachment-1", Kind = "attachment", FileName = "notes.txt", MediaType = "text/plain", LengthBytes = 12 }
            },
            Diagnostics = new[] {
                new OfficeDocumentDiagnostic {
                    Severity = OfficeDocumentDiagnosticSeverity.Warning,
                    Code = "source-warning",
                    Message = "Source normalization warning",
                    Source = "Email"
                }
            }
        };

        PdfDocumentConversionResult result = source.ToPdfDocumentResult(new ReaderPdfProjectionOptions {
            AssetPolicy = ReaderPdfAssetPolicy.ListMetadata
        });

        string text = PdfReadDocument.Open(result.ToBytes()).ExtractText();
        Assert.Contains("The release is ready.", text, StringComparison.Ordinal);
        Assert.Contains("notes.txt", text, StringComparison.Ordinal);
        Assert.Contains(result.Warnings, warning => warning.Code == "source-warning");
        Assert.Contains(result.Warnings, warning => warning.Code == "reader-email-policy" && warning.Severity == PdfConversionWarningSeverity.Information);
    }

    [Fact]
    public void ReaderProjection_ReconcilesPageAndDocumentContentWithoutDuplicatingTablesOrLosingResources() {
        var tableLocation = new ReaderLocation {
            Path = "chapter.xhtml",
            SourceBlockIndex = 1,
            SourceBlockKind = "table",
            BlockAnchor = "table-1",
            TableIndex = 0
        };
        var source = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Epub,
            Pages = new[] {
                new OfficeDocumentPage {
                    Name = "Chapter",
                    Blocks = new[] {
                        new OfficeDocumentBlock { Id = "before", Kind = "paragraph", Text = "BEFOREMARKER", Location = new ReaderLocation { Path = "chapter.xhtml", SourceBlockIndex = 0 } },
                        new OfficeDocumentBlock { Id = "table-1", Kind = "table", Text = "DUPLICATETABLEMARKER", Location = tableLocation },
                        new OfficeDocumentBlock { Id = "after", Kind = "paragraph", Text = "AFTERMARKER", Location = new ReaderLocation { Path = "chapter.xhtml", SourceBlockIndex = 2 } }
                    },
                    Tables = new[] {
                        new ReaderTable {
                            Location = tableLocation,
                            Columns = new[] { "Column" },
                            Rows = new[] { (IReadOnlyList<string>)new[] { "TABLECELLMARKER" } },
                            TotalRowCount = 1
                        }
                    }
                }
            },
            Blocks = new[] {
                new OfficeDocumentBlock { Id = "before", Kind = "paragraph", Text = "BEFOREMARKER" },
                new OfficeDocumentBlock { Id = "resource-note", Kind = "paragraph", Text = "DOCUMENTMARKER" }
            },
            Assets = new[] {
                new OfficeDocumentAsset { Id = "package-resource", Kind = "resource", FileName = "styles.css", MediaType = "text/css" }
            }
        };

        PdfDocumentConversionResult result = source.ToPdfDocumentResult(new ReaderPdfProjectionOptions {
            AssetPolicy = ReaderPdfAssetPolicy.ListMetadata
        });

        string text = PdfReadDocument.Open(result.ToBytes()).ExtractText();
        Assert.DoesNotContain("DUPLICATETABLEMARKER", text, StringComparison.Ordinal);
        Assert.Equal(1, CountOccurrences(text, "BEFOREMARKER"));
        Assert.Equal(1, CountOccurrences(text, "TABLECELLMARKER"));
        Assert.True(text.IndexOf("BEFOREMARKER", StringComparison.Ordinal) < text.IndexOf("TABLECELLMARKER", StringComparison.Ordinal));
        Assert.True(text.IndexOf("TABLECELLMARKER", StringComparison.Ordinal) < text.IndexOf("AFTERMARKER", StringComparison.Ordinal));
        Assert.Contains("DOCUMENTMARKER", text, StringComparison.Ordinal);
        Assert.Contains("styles.css", text, StringComparison.Ordinal);
    }

    [Fact]
    public void ReaderProjection_UsesPageFallbackToDeduplicateLocationlessTableClone() {
        var aggregate = new ReaderTable {
            Title = "Settings",
            Location = new ReaderLocation { Path = "settings.pdf", Page = 7, TableIndex = 0 },
            Columns = new[] { "Key", "Value" },
            Rows = new[] { (IReadOnlyList<string>)new[] { "Mode", "SAFEMARKER" } },
            TotalRowCount = 1
        };
        var pageClone = new ReaderTable {
            Title = aggregate.Title,
            Columns = aggregate.Columns,
            Rows = aggregate.Rows,
            TotalRowCount = aggregate.TotalRowCount
        };
        var source = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Pdf,
            Tables = new[] { aggregate },
            Pages = new[] {
                new OfficeDocumentPage {
                    Number = 7,
                    Location = new ReaderLocation { Path = "settings.pdf" },
                    Tables = new[] { pageClone }
                }
            }
        };

        string text = PdfReadDocument.Open(source.ToPdfDocumentResult().ToBytes()).ExtractText();

        Assert.Equal(1, CountOccurrences(text, "SAFEMARKER"));
    }

    [Fact]
    public void ReaderProjection_ReportsSelectedAnimationFrameAndHonorsExactPolicy() {
        byte[] animatedGif = CreateTwoFrameGif();
        var source = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Unknown,
            Assets = new[] {
                new OfficeDocumentAsset { Id = "animation", Kind = "image", FileName = "animation.gif", MediaType = "image/gif", PayloadBytes = animatedGif }
            }
        };

        PdfDocumentConversionResult selected = source.ToPdfDocumentResult();
        PdfDocumentConversionResult rejected = source.ToPdfDocumentResult(new ReaderPdfProjectionOptions {
            RasterDecodeOptions = new OfficeRasterDecodeOptions {
                AnimationPolicy = OfficeRasterAnimationPolicy.RejectAnimated
            }
        });

        Assert.Contains(selected.Warnings, warning => warning.Code == "reader-asset-animation-frame-selected");
        Assert.Contains(rejected.Warnings, warning => warning.Code == "reader-asset-animation-rejected");
        Assert.Contains("animation.gif", PdfReadDocument.Open(rejected.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void ReaderProjection_VisioEvidenceReflectsListPolicyInsteadOfClaimingEmbedding() {
        var source = new OfficeDocumentReadResult {
            Kind = ReaderInputKind.Visio,
            Assets = new[] {
                new OfficeDocumentAsset {
                    Id = "preview",
                    Kind = "preview-image",
                    FileName = "preview.gif",
                    MediaType = "image/gif",
                    PayloadBytes = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==")
                }
            }
        };

        PdfDocumentConversionResult result = source.ToPdfDocumentResult(new ReaderPdfProjectionOptions {
            AssetPolicy = ReaderPdfAssetPolicy.ListMetadata
        });

        Assert.Contains(result.Warnings, warning => warning.Code == "reader-visio-preview-listed");
        Assert.DoesNotContain(result.Warnings, warning => warning.Code == "reader-visio-preview-embedded");
    }

    private static byte[] CreateTwoFrameGif() {
        byte[] first = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");
        const int imageDescriptorOffset = 19;
        return first.Take(first.Length - 1)
            .Concat(first.Skip(imageDescriptorOffset).Take(first.Length - imageDescriptorOffset - 1))
            .Concat(new byte[] { 0x3B })
            .ToArray();
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

    private sealed class PageAwareComponent : IPdfContextComponent {
        internal int Invocations { get; private set; }

        public void Compose(PdfItemCompose content, PdfFlowContext context) {
            Invocations++;
            content.Paragraph(paragraph => paragraph.Text("Context page " + context.PageNumber));
        }
    }
}
