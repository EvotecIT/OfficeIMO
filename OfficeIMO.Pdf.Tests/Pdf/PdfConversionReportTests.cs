using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConversionReportTests {
    [Fact]
    public void PdfConversionReport_SummarizeGroupsWarningsForProofAndWrapperRouting() {
        var report = new PdfConversionReport();
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Word.Pdf",
            "UnsupportedShape",
            "word:shape[1]",
            "Shape was simplified."));
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Word.Pdf",
            "UnsupportedShape",
            "word:shape[2]",
            "Shape was simplified."));
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Excel.Pdf",
            "FormulaValueMissing",
            "sheet:Summary!C4",
            "Formula value was unavailable.",
            PdfConversionWarningSeverity.Error));
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Pdf",
            "ConversionContext",
            string.Empty,
            "Conversion used fallback profile.",
            PdfConversionWarningSeverity.Information));

        PdfConversionReportSummary summary = report.Summarize();

        Assert.True(summary.HasWarnings);
        Assert.True(summary.HasErrors);
        Assert.Equal(4, summary.TotalCount);
        Assert.Equal(1, summary.InformationCount);
        Assert.Equal(2, summary.WarningCount);
        Assert.Equal(1, summary.ErrorCount);
        Assert.Equal(2, summary.ConverterCounts["OfficeIMO.Word.Pdf"]);
        Assert.Equal(1, summary.ConverterCounts["OfficeIMO.Excel.Pdf"]);
        Assert.Equal(2, summary.CodeCounts["UnsupportedShape"]);
        Assert.Equal(1, summary.CodeCounts["FormulaValueMissing"]);
        Assert.Equal(1, summary.SourceCounts["word:shape[1]"]);
        Assert.False(summary.SourceCounts.ContainsKey(string.Empty));
    }

    [Fact]
    public void PdfConversionReport_RequireNoWarningsReturnsReportWhenClean() {
        var report = new PdfConversionReport();

        PdfConversionReport returned = report.RequireNoWarnings();

        Assert.Same(report, returned);
        Assert.Same(report, report.RequireNoErrorWarnings());
        Assert.False(report.HasErrors);
    }

    [Fact]
    public void PdfConversionReport_RequireNoWarningsFailsOnAnySeverity() {
        var report = new PdfConversionReport();
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Tests",
            "DecorativeFallback",
            "test",
            "Decorative content was simplified.",
            PdfConversionWarningSeverity.Information));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => report.RequireNoWarnings());

        Assert.Contains("PDF conversion produced warnings.", exception.Message, StringComparison.Ordinal);
        Assert.Contains("DecorativeFallback", exception.Message, StringComparison.Ordinal);
        Assert.False(report.HasErrors);
        Assert.Same(report, report.RequireNoErrorWarnings());
    }

    [Fact]
    public void PdfConversionReport_RequireNoErrorWarningsFailsOnlyOnErrors() {
        var report = new PdfConversionReport();
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Tests",
            "FormulaValueMissing",
            "test",
            "Formula value was unavailable.",
            PdfConversionWarningSeverity.Error));

        Assert.True(report.HasErrors);
        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() => report.RequireNoErrorWarnings());

        Assert.Contains("PDF conversion produced error warnings.", exception.Message, StringComparison.Ordinal);
        Assert.Contains("FormulaValueMissing", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocumentConversionResult_SummaryUsesCapturedConversionReportSnapshot() {
        var report = new PdfConversionReport();
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Html.Pdf",
            "StylesheetResourceRejectedByPolicy",
            "html:head/link[1]",
            "Stylesheet was blocked by the configured resource policy."));

        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Summary snapshot proof")),
            report);

        report.Clear();
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Tests",
            "LaterWarning",
            "test",
            "This warning belongs to a later conversion."));

        PdfConversionReportSummary summary = result.Summary;

        Assert.Equal(1, summary.TotalCount);
        Assert.Equal(1, summary.WarningCount);
        Assert.Equal(1, summary.ConverterCounts["OfficeIMO.Html.Pdf"]);
        Assert.Equal(1, summary.CodeCounts["StylesheetResourceRejectedByPolicy"]);
        Assert.False(summary.CodeCounts.ContainsKey("LaterWarning"));
        Assert.Contains("Summary snapshot proof", PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocumentConversionResult_WithDocumentRefreshesFromOriginalConversionReport() {
        var report = new PdfConversionReport();
        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Refresh source proof")),
            report);

        PdfDocumentConversionResult processed = result.WithValue(result.Value.UpdateMetadata(title: "Processed refresh source proof"));
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Tests",
            "SaveTimeWarning",
            "save:document",
            "Warning emitted after post-processing."));

        processed.ToBytes();

        PdfConversionWarning warning = Assert.Single(processed.Warnings);
        Assert.Equal("SaveTimeWarning", warning.Code);
        Assert.DoesNotContain(result.Warnings, item => item.Code == "SaveTimeWarning");
    }

    [Fact]
    public void PdfDocumentConversionResult_FailedSerializationPreservesRecordedDiagnostics() {
        var report = new PdfConversionReport();
        var options = new PdfOptions().ReportDiagnosticsTo(report, "OfficeIMO.Tests");
        var result = new PdfDocumentConversionResult(
            PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("مرحبا")),
            report);

        Assert.ThrowsAny<ArgumentException>(() => result.ToBytes());

        Assert.Contains(result.Warnings, warning =>
            warning.Code == "unsupported-bidirectional-text-layout" &&
            warning.Converter == "OfficeIMO.Tests");
        Assert.Contains(result.Warnings, warning =>
            warning.Code == "unsupported-complex-script-shaping" &&
            warning.Converter == "OfficeIMO.Tests");
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofCapturesTextWarningsAndProcessedDocument() {
        var report = new PdfConversionReport();
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Markdown.Pdf",
            "AcceptedTableSimplification",
            "markdown:table[1]",
            "Table border details were simplified."));

        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Conversion proof text")),
            report);

        PdfDocumentConversionResult processed = result.Process(document => document.UpdateMetadata(title: "Proof metadata"));
        PdfConversionProofReport proof = processed.AssessProof(new PdfConversionProofOptions()
            .RequireTextMarkers("Conversion proof text")
            .RequireWarningCodes("AcceptedTableSimplification")
            .RequireWarningSources("markdown:table[1]")
            .RequireNoErrors());

        Assert.True(proof.IsSatisfied, proof.Summary);
        Assert.Empty(proof.Issues);
        Assert.NotNull(proof.DocumentInfo);
        Assert.Equal("Proof metadata", proof.DocumentInfo!.Metadata.Title);
        Assert.Contains("Conversion proof text", proof.ExtractedText, StringComparison.Ordinal);
        Assert.Equal(1, proof.WarningSummary.TotalCount);
        Assert.Equal(1, proof.WarningSummary.CodeCounts["AcceptedTableSimplification"]);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofCapturesArtifactHash() {
        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Artifact hash proof")),
            new PdfConversionReport());

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .RequireTextMarkers("Artifact hash proof"));

        Assert.True(proof.IsSatisfied, proof.Summary);
        Assert.True(proof.ArtifactByteCount > 0);
        Assert.Equal(64, proof.ArtifactSha256.Length);
        Assert.Matches("^[0-9a-f]{64}$", proof.ArtifactSha256);

        PdfConversionProofReport pinnedProof = result.AssessProof(new PdfConversionProofOptions()
            .RequireArtifactSha256(proof.ArtifactSha256));

        Assert.True(pinnedProof.IsSatisfied, pinnedProof.Summary);
        Assert.Equal(proof.ArtifactSha256, pinnedProof.ArtifactSha256);
        Assert.Equal(proof.ArtifactByteCount, pinnedProof.ArtifactByteCount);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofCapturesRequiredPageCount() {
        var result = new PdfDocumentConversionResult(
            PdfDocument.Create()
                .Paragraph(paragraph => paragraph.Text("First page proof"))
                .PageBreak()
                .Paragraph(paragraph => paragraph.Text("Second page proof")),
            new PdfConversionReport());

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .RequirePageCount(2)
            .RequireTextMarkers("First page proof", "Second page proof"));

        Assert.True(proof.IsSatisfied, proof.Summary);
        Assert.Equal(2, proof.DocumentInfo!.PageCount);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofCapturesRequiredPageSize() {
        var result = new PdfDocumentConversionResult(
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 320,
                    PageHeight = 240
                })
                .Paragraph(paragraph => paragraph.Text("Page geometry proof")),
            new PdfConversionReport());

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .RequirePageSize(320, 240)
            .RequireLogicalSignals("page-geometry"));

        Assert.True(proof.IsSatisfied, proof.Summary);
        Assert.Equal(320, proof.DocumentInfo!.Pages[0].Width);
        Assert.Equal(240, proof.DocumentInfo.Pages[0].Height);
        Assert.Contains("page-geometry", proof.LogicalSignals);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofCapturesRequiredMetadata() {
        var result = new PdfDocumentConversionResult(
            PdfDocument.Create()
                .Meta(
                    title: "Metadata proof title",
                    author: "OfficeIMO",
                    subject: "PDF proof contract",
                    keywords: "pdf, metadata, proof")
                .Paragraph(paragraph => paragraph.Text("Metadata proof body")),
            new PdfConversionReport());

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .RequireMetadata(
                title: "Metadata proof title",
                author: "OfficeIMO",
                subject: "PDF proof contract",
                keywords: "pdf, metadata, proof")
            .RequireLogicalSignals("metadata", "document-metadata"));

        Assert.True(proof.IsSatisfied, proof.Summary);
        Assert.Equal("Metadata proof title", proof.DocumentInfo!.Metadata.Title);
        Assert.Equal("OfficeIMO", proof.DocumentInfo.Metadata.Author);
        Assert.Equal("PDF proof contract", proof.DocumentInfo.Metadata.Subject);
        Assert.Equal("pdf, metadata, proof", proof.DocumentInfo.Metadata.Keywords);
        Assert.Contains("metadata", proof.LogicalSignals);
        Assert.Contains("document-metadata", proof.LogicalSignals);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofReportsPageCountMismatch() {
        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Single page proof")),
            new PdfConversionReport());

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .RequirePageCount(2));

        Assert.False(proof.IsSatisfied);
        PdfConversionProofIssue issue = Assert.Single(proof.Issues, item => item.Feature == "PageCount");
        Assert.Equal("2", issue.Expected);
        Assert.Equal("1", issue.Actual);
        Assert.Contains("PageCount", proof.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofReportsPageSizeMismatch() {
        var result = new PdfDocumentConversionResult(
            PdfDocument.Create(new PdfOptions {
                    PageWidth = 320,
                    PageHeight = 240
                })
                .Paragraph(paragraph => paragraph.Text("Page size mismatch proof")),
            new PdfConversionReport());

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .RequirePageSize(320, 260));

        Assert.False(proof.IsSatisfied);
        PdfConversionProofIssue issue = Assert.Single(proof.Issues, item => item.Feature == "PageSize");
        Assert.Equal("320x260", issue.Expected);
        Assert.Equal("page 1 320x240", issue.Actual);
        Assert.Contains("PageSize", proof.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofReportsMetadataMismatch() {
        var result = new PdfDocumentConversionResult(
            PdfDocument.Create()
                .Meta(title: "Actual proof title", author: "OfficeIMO")
                .Paragraph(paragraph => paragraph.Text("Metadata mismatch proof")),
            new PdfConversionReport());

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .RequireMetadata(title: "Expected proof title"));

        Assert.False(proof.IsSatisfied);
        PdfConversionProofIssue issue = Assert.Single(proof.Issues, item => item.Feature == "Metadata.Title");
        Assert.Equal("Expected proof title", issue.Expected);
        Assert.Equal("Actual proof title", issue.Actual);
        Assert.Contains("Metadata.Title", proof.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofReportsArtifactHashMismatch() {
        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Artifact mismatch proof")),
            new PdfConversionReport());

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .RequireArtifactSha256(new string('0', 64)));

        Assert.False(proof.IsSatisfied);
        PdfConversionProofIssue issue = Assert.Single(proof.Issues, item => item.Feature == "ArtifactSha256");
        Assert.Equal(new string('0', 64), issue.Expected);
        Assert.Equal(proof.ArtifactSha256, issue.Actual);
        Assert.Contains("ArtifactSha256", proof.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofAcceptsDeclaredWarningCodes() {
        var report = new PdfConversionReport();
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Html.Pdf",
            "StylesheetResourceRejectedByPolicy",
            "html:head/link[1]",
            "Stylesheet was blocked by the configured resource policy."));
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Html.Pdf",
            "UnsupportedCssDeclaration",
            "html:style[1]",
            "CSS declaration was ignored."));

        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Accepted degradation proof")),
            report);

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .RequireTextMarkers("Accepted degradation proof")
            .RequireWarningCodes("StylesheetResourceRejectedByPolicy", "UnsupportedCssDeclaration")
            .AcceptWarningCodes("StylesheetResourceRejectedByPolicy", "UnsupportedCssDeclaration")
            .RequireNoUnexpectedWarningCodes());

        Assert.True(proof.IsSatisfied, proof.Summary);
        Assert.Empty(proof.Issues);
        Assert.Equal(2, proof.WarningSummary.TotalCount);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofReportsUnexpectedWarningCodes() {
        var report = new PdfConversionReport();
        report.Add(new PdfConversionWarning(
            "OfficeIMO.PowerPoint.Pdf",
            "AcceptedMediaPlaceholder",
            "slide:1/media[1]",
            "Media placeholder was represented as a poster frame."));
        report.Add(new PdfConversionWarning(
            "OfficeIMO.PowerPoint.Pdf",
            "UnexpectedSmartArtFallback",
            "slide:1/smartArt[1]",
            "SmartArt fallback was not declared by the proof contract."));

        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Unexpected warning proof")),
            report);

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .AcceptWarningCodes("AcceptedMediaPlaceholder")
            .RequireNoUnexpectedWarningCodes());

        Assert.False(proof.IsSatisfied);
        PdfConversionProofIssue issue = Assert.Single(proof.Issues, item => item.Feature == "UnexpectedWarningCode");
        Assert.Equal("UnexpectedSmartArtFallback", issue.Actual);
        Assert.Contains("UnexpectedSmartArtFallback", proof.Summary, StringComparison.Ordinal);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofCapturesLogicalSignals() {
        var report = new PdfConversionReport();
        var options = new PdfOptions {
            CreateOutlineFromHeadings = true,
            IncludePageLabels = true,
            PageNumberStyle = PdfPageNumberStyle.UpperRoman,
            PageNumberStart = 3,
            PageLabelPrefix = "A-",
            ViewerPreferences = new PdfViewerPreferencesOptions {
                DisplayDocTitle = true,
                HideToolbar = true
            }
        };
        PdfDocument document = PdfDocument.Create(options)
            .PdfAIdentification(3, "B")
            .ConfigurePdfUaGroundwork("en-GB")
            .Meta(title: "Logical proof metadata", author: "OfficeIMO", subject: "PDF proof contract", keywords: "pdf, metadata, proof")
            .Language("en-GB")
            .CatalogView(PdfCatalogPageMode.UseOutlines, PdfCatalogPageLayout.SinglePage)
            .OpenAction(1, 640D, PdfOpenActionDestinationMode.FitHorizontal)
            .SrgbOutputIntent()
            .AttachFile("proof.xml", new byte[] { 60, 112, 114, 111, 111, 102, 32, 47, 62 }, "application/xml", PdfAssociatedFileRelationship.Data, "Proof XML")
            .Bookmark("LogicalAnchor")
            .H1("Logical proof heading")
            .Paragraph(paragraph => paragraph
                .Text("Logical proof body with ")
                .Link("support link", "https://evotec.xyz/support"))
            .TextField("Customer.Name", value: "Ada")
            .Table(new[] {
                new[] { "Metric", "Value" },
                new[] { "Quality", "Premium" }
            });
        var result = new PdfDocumentConversionResult(document, report);

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .RequireTextMarkers("Logical proof heading", "Premium")
            .RequireOutlineTitles("Logical proof heading")
            .RequireLinkUris("https://evotec.xyz/support")
            .RequireFormFieldNames("Customer.Name")
            .RequireNamedDestinationNames("LogicalAnchor")
            .RequirePageLabelRange(1, PdfPageNumberStyle.UpperRoman, 3, "A-")
            .RequireAttachmentFileNames("proof.xml")
            .RequireOutputIntentSubtypes("GTS_PDFA1")
            .RequireOutputConditionIdentifiers(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier)
            .RequireCatalogLanguage("en-GB")
            .RequireCatalogView(PdfCatalogPageMode.UseOutlines, PdfCatalogPageLayout.SinglePage)
            .RequireOpenAction(1, PdfOpenActionDestinationMode.FitHorizontal)
            .RequireViewerPreference("DisplayDocTitle", true)
            .RequireViewerPreference("HideToolbar", true)
            .RequireXmpMetadata(
                title: "Logical proof metadata",
                creator: "OfficeIMO",
                producer: "OfficeIMO.Pdf",
                keywords: "pdf, metadata, proof")
            .RequireXmpSubjects("pdf", "metadata", "proof")
            .RequireXmpPdfAIdentification(3, "B")
            .RequireXmpPdfUaIdentification(1)
            .RequireTaggedStructureTypes("Document", "H1", "P")
            .RequireTaggedStructureElementCountAtLeast(3)
            .RequireTaggedMarkedContentReferencesAtLeast(1)
            .RequireLogicalSignals("page-count", "page-geometry", "metadata", "text", "headings", "outlines", "links", "form-fields", "tables", "named-destinations", "page-labels", "attachments", "output-intents", "catalog-view", "open-action", "viewer-preferences", "xmp", "tagged-content"));

        Assert.True(proof.IsSatisfied, proof.Summary);
        Assert.NotNull(proof.LogicalDocument);
        Assert.Contains(proof.DocumentInfo!.Outlines, outline => outline.Title == "Logical proof heading");
        Assert.Contains("https://evotec.xyz/support", proof.DocumentInfo.LinkUris);
        Assert.Contains("Customer.Name", proof.DocumentInfo.FormFieldNames);
        Assert.Contains("LogicalAnchor", proof.DocumentInfo.NamedDestinationNames);
        Assert.Contains("proof.xml", proof.DocumentInfo.AttachmentFileNames);
        Assert.Contains("GTS_PDFA1", proof.DocumentInfo.OutputIntentSubtypes);
        Assert.Contains(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier, proof.DocumentInfo.OutputConditionIdentifiers);
        Assert.Equal("en-GB", proof.DocumentInfo.CatalogLanguage);
        Assert.Equal("UseOutlines", proof.DocumentInfo.CatalogPageMode);
        Assert.Equal("SinglePage", proof.DocumentInfo.CatalogPageLayout);
        Assert.Equal(1, proof.DocumentInfo.OpenAction!.PageNumber);
        Assert.Equal(PdfOpenActionDestinationMode.FitHorizontal, proof.DocumentInfo.OpenAction.DestinationMode);
        Assert.True(proof.DocumentInfo.ViewerPreferences!.GetBoolean("DisplayDocTitle"));
        Assert.True(proof.DocumentInfo.ViewerPreferences.GetBoolean("HideToolbar"));
        Assert.Equal("Logical proof metadata", proof.DocumentInfo.XmpMetadata!.Title);
        Assert.Equal("OfficeIMO", proof.DocumentInfo.XmpMetadata.Creator);
        Assert.Contains("metadata", proof.DocumentInfo.XmpMetadata.Subjects);
        Assert.Equal(3, proof.DocumentInfo.XmpMetadata.PdfAPart);
        Assert.Equal("B", proof.DocumentInfo.XmpMetadata.PdfAConformance);
        Assert.Equal(1, proof.DocumentInfo.XmpMetadata.PdfUaPart);
        Assert.Contains("Document", proof.DocumentInfo.TaggedContent!.StructureTypes);
        Assert.Contains("H1", proof.DocumentInfo.TaggedContent.StructureTypes);
        Assert.Contains("P", proof.DocumentInfo.TaggedContent.StructureTypes);
        Assert.True(proof.DocumentInfo.TaggedContent.StructureElementCount >= 3);
        Assert.True(proof.DocumentInfo.TaggedContent.MarkedContentReferenceCount >= 1);
        PdfPageLabel label = Assert.Single(proof.DocumentInfo.PageLabels);
        Assert.Equal(1, label.StartPageNumber);
        Assert.Equal("R", label.Style);
        Assert.Equal("A-", label.Prefix);
        Assert.Equal(3, label.StartNumber);
        Assert.Contains("page-geometry", proof.LogicalSignals);
        Assert.Contains("metadata", proof.LogicalSignals);
        Assert.Contains("headings", proof.LogicalSignals);
        Assert.Contains("outlines", proof.LogicalSignals);
        Assert.Contains("links", proof.LogicalSignals);
        Assert.Contains("form-fields", proof.LogicalSignals);
        Assert.Contains("tables", proof.LogicalSignals);
        Assert.Contains("named-destinations", proof.LogicalSignals);
        Assert.Contains("page-labels", proof.LogicalSignals);
        Assert.Contains("attachments", proof.LogicalSignals);
        Assert.Contains("output-intents", proof.LogicalSignals);
        Assert.Contains("catalog-view", proof.LogicalSignals);
        Assert.Contains("open-action", proof.LogicalSignals);
        Assert.Contains("viewer-preferences", proof.LogicalSignals);
        Assert.Contains("xmp", proof.LogicalSignals);
        Assert.Contains("tagged-content", proof.LogicalSignals);
        Assert.NotEmpty(proof.LogicalDocument!.Outlines);
        Assert.NotEmpty(proof.LogicalDocument.Links);
        Assert.NotEmpty(proof.LogicalDocument.FormFields);
        Assert.NotEmpty(proof.LogicalDocument.Tables);
        Assert.NotEmpty(proof.LogicalDocument.NamedDestinations);
        Assert.NotEmpty(proof.LogicalDocument.PageLabels);
        Assert.NotEmpty(proof.LogicalDocument.Attachments);
        Assert.NotEmpty(proof.LogicalDocument.OutputIntents);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofCapturesOptionalContentLayers() {
        PdfDocument document = PdfDocument.Open(PdfOptionalContentSupport.BuildOptionalContentMetadataPdf());
        var result = new PdfDocumentConversionResult(document, new PdfConversionReport());

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .RequireOptionalContentGroupCountAtLeast(2)
            .RequireOptionalContentGroupNames("Print layer", "Hidden layer")
            .RequireOptionalContentDefaultConfiguration("Default layers", "OfficeIMO fixture", "ON")
            .RequireOptionalContentVisibleGroupNames("Print layer")
            .RequireOptionalContentHiddenGroupNames("Hidden layer")
            .RequireOptionalContentLockedGroupNames("Hidden layer")
            .RequireOptionalContentOrderedGroupNames("Print layer", "Hidden layer")
            .RequireLogicalSignals("optional-content", "layers"));

        Assert.True(proof.IsSatisfied, proof.Summary);
        Assert.Equal(2, proof.DocumentInfo!.OptionalContentGroupCount);
        Assert.Equal(new[] { "Print layer", "Hidden layer" }, proof.DocumentInfo.OptionalContentGroupNames);
        Assert.Equal("Default layers", proof.DocumentInfo.OptionalContent!.DefaultConfigurationName);
        Assert.Equal("OfficeIMO fixture", proof.DocumentInfo.OptionalContent.DefaultConfigurationCreator);
        Assert.Equal("ON", proof.DocumentInfo.OptionalContent.BaseState);
        PdfOptionalContentGroup printLayer = Assert.Single(proof.DocumentInfo.GetOptionalContentGroupsByName("Print layer"));
        Assert.True(printLayer.IsInitiallyVisible);
        Assert.False(printLayer.IsLocked);
        PdfOptionalContentGroup hiddenLayer = Assert.Single(proof.DocumentInfo.GetOptionalContentGroupsByName("Hidden layer"));
        Assert.False(hiddenLayer.IsInitiallyVisible);
        Assert.True(hiddenLayer.IsLocked);
        Assert.Contains("optional-content", proof.LogicalSignals);
        Assert.Contains("layers", proof.LogicalSignals);
    }

    [Fact]
    public void PdfDocumentConversionResult_AssessProofReportsMissingEvidenceAndErrorWarnings() {
        var report = new PdfConversionReport();
        report.Add(new PdfConversionWarning(
            "OfficeIMO.Excel.Pdf",
            "FormulaValueMissing",
            "sheet:Summary!C4",
            "Formula value was unavailable.",
            PdfConversionWarningSeverity.Error));

        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Available proof text")),
            report);

        PdfConversionProofReport proof = result.AssessProof(new PdfConversionProofOptions()
            .RequireTextMarkers("Missing proof text")
            .RequireLogicalSignals("links")
            .RequireOutlineTitles("Missing outline")
            .RequireLinkUris("https://evotec.xyz/missing")
            .RequireFormFieldNames("Missing.Field")
            .RequireNamedDestinationNames("MissingAnchor")
            .RequirePageLabelRange(1, PdfPageNumberStyle.LowerRoman, 1, "front-")
            .RequireAttachmentFileNames("missing.xml")
            .RequireOutputIntentSubtypes("GTS_PDFA1")
            .RequireOutputConditionIdentifiers(PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier)
            .RequireOptionalContentGroupCountAtLeast(1)
            .RequireOptionalContentGroupNames("Missing layer")
            .RequireOptionalContentDefaultConfiguration("Missing layers", "OfficeIMO fixture", "ON")
            .RequireOptionalContentVisibleGroupNames("Visible layer")
            .RequireOptionalContentHiddenGroupNames("Hidden layer")
            .RequireOptionalContentLockedGroupNames("Locked layer")
            .RequireOptionalContentOrderedGroupNames("Ordered layer")
            .RequireCatalogLanguage("en-GB")
            .RequireCatalogView(PdfCatalogPageMode.UseOutlines, PdfCatalogPageLayout.SinglePage)
            .RequireOpenAction(1, PdfOpenActionDestinationMode.FitHorizontal)
            .RequireViewerPreference("DisplayDocTitle", true)
            .RequireXmpMetadata(title: "Missing XMP title", creator: "OfficeIMO", producer: "OfficeIMO.Pdf")
            .RequireXmpSubjects("missing-subject")
            .RequireXmpPdfAIdentification(3, "B")
            .RequireXmpPdfUaIdentification(1)
            .RequireTaggedStructureTypes("Document")
            .RequireTaggedStructureElementCountAtLeast(1)
            .RequireTaggedMarkedContentReferencesAtLeast(1)
            .RequireWarningCodes("ExpectedWarning")
            .RequireWarningSources("sheet:Missing!A1")
            .RequireNoErrors());

        Assert.False(proof.IsSatisfied);
        Assert.Contains(proof.Issues, issue => issue.Feature == "TextMarker" && issue.Expected == "Missing proof text");
        Assert.Contains(proof.Issues, issue => issue.Feature == "LogicalSignal" && issue.Expected == "links");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OutlineTitle" && issue.Expected == "Missing outline");
        Assert.Contains(proof.Issues, issue => issue.Feature == "LinkUri" && issue.Expected == "https://evotec.xyz/missing");
        Assert.Contains(proof.Issues, issue => issue.Feature == "FormFieldName" && issue.Expected == "Missing.Field");
        Assert.Contains(proof.Issues, issue => issue.Feature == "NamedDestination" && issue.Expected == "MissingAnchor");
        Assert.Contains(proof.Issues, issue => issue.Feature == "PageLabel" && issue.Expected == "page 1 r start 1 prefix front-");
        Assert.Contains(proof.Issues, issue => issue.Feature == "AttachmentFileName" && issue.Expected == "missing.xml");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OutputIntentSubtype" && issue.Expected == "GTS_PDFA1");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OutputConditionIdentifier" && issue.Expected == PdfIccProfiles.SrgbIec6196621OutputConditionIdentifier);
        Assert.Contains(proof.Issues, issue => issue.Feature == "OptionalContent.GroupCount" && issue.Expected == "at least 1");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OptionalContent.GroupName" && issue.Expected == "Missing layer");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OptionalContent.DefaultConfigurationName" && issue.Expected == "Missing layers");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OptionalContent.DefaultConfigurationCreator" && issue.Expected == "OfficeIMO fixture");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OptionalContent.BaseState" && issue.Expected == "ON");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OptionalContent.VisibleGroupName" && issue.Expected == "Visible layer");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OptionalContent.HiddenGroupName" && issue.Expected == "Hidden layer");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OptionalContent.LockedGroupName" && issue.Expected == "Locked layer");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OptionalContent.OrderedGroupName" && issue.Expected == "Ordered layer");
        Assert.Contains(proof.Issues, issue => issue.Feature == "CatalogLanguage" && issue.Expected == "en-GB");
        Assert.Contains(proof.Issues, issue => issue.Feature == "CatalogPageMode" && issue.Expected == "UseOutlines");
        Assert.Contains(proof.Issues, issue => issue.Feature == "CatalogPageLayout" && issue.Expected == "SinglePage");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OpenAction.PageNumber" && issue.Expected == "1");
        Assert.Contains(proof.Issues, issue => issue.Feature == "OpenAction.DestinationMode" && issue.Expected == "FitHorizontal");
        Assert.Contains(proof.Issues, issue => issue.Feature == "ViewerPreference.DisplayDocTitle" && issue.Expected == "true");
        Assert.Contains(proof.Issues, issue => issue.Feature == "Xmp.Title" && issue.Expected == "Missing XMP title");
        Assert.Contains(proof.Issues, issue => issue.Feature == "Xmp.Creator" && issue.Expected == "OfficeIMO");
        Assert.Contains(proof.Issues, issue => issue.Feature == "Xmp.Producer" && issue.Expected == "OfficeIMO.Pdf");
        Assert.Contains(proof.Issues, issue => issue.Feature == "Xmp.Subject" && issue.Expected == "missing-subject");
        Assert.Contains(proof.Issues, issue => issue.Feature == "Xmp.PdfAPart" && issue.Expected == "3");
        Assert.Contains(proof.Issues, issue => issue.Feature == "Xmp.PdfAConformance" && issue.Expected == "B");
        Assert.Contains(proof.Issues, issue => issue.Feature == "Xmp.PdfUaPart" && issue.Expected == "1");
        Assert.Contains(proof.Issues, issue => issue.Feature == "TaggedContent.StructureType" && issue.Expected == "Document");
        Assert.Contains(proof.Issues, issue => issue.Feature == "TaggedContent.StructureElementCount" && issue.Expected == "at least 1");
        Assert.Contains(proof.Issues, issue => issue.Feature == "TaggedContent.MarkedContentReferenceCount" && issue.Expected == "at least 1");
        Assert.Contains(proof.Issues, issue => issue.Feature == "WarningCode" && issue.Expected == "ExpectedWarning");
        Assert.Contains(proof.Issues, issue => issue.Feature == "WarningSource" && issue.Expected == "sheet:Missing!A1");
        Assert.Contains(proof.Issues, issue => issue.Feature == "WarningSeverity" && issue.Actual == "1");
        Assert.Contains("PDF conversion proof failed", proof.Summary, StringComparison.Ordinal);
        Assert.Throws<InvalidOperationException>(() => result.AssertProof(new PdfConversionProofOptions().RequireTextMarkers("Missing proof text")));
    }
}
