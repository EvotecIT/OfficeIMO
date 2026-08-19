using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using OfficeIMO.Rtf;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void OfficeHtmlExportOptions_UseSourceSpecificProfilesAndRejectCrossFormatCompatibilityValues() {
        Assert.Equal(ExcelHtmlExportProfile.SemanticTables,
            ExcelHtmlSaveOptions.CreateSemanticTablesProfile().ExportProfile);
        Assert.Equal(ExcelHtmlExportProfile.VisualReview,
            ExcelHtmlSaveOptions.CreateVisualReviewProfile().ExportProfile);
        Assert.Equal(PowerPointHtmlExportProfile.SemanticSlides,
            PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile().ExportProfile);
        Assert.Equal(PowerPointHtmlExportProfile.VisualReview,
            PowerPointHtmlSaveOptions.CreateVisualReviewProfile().ExportProfile);
        Assert.Equal(WordHtmlExportProfile.PrintReview,
            WordToHtmlOptions.CreatePrintReviewProfile().ExportProfile);
        Assert.Equal(RtfHtmlExportProfile.PrintReview,
            RtfToHtmlOptions.CreatePrintReviewProfile().ExportProfile);
        Assert.Equal(HtmlConversionProfile.Semantic,
            ExcelHtmlSaveOptions.CreateSemanticTablesProfile().SharedProfile);
        Assert.Equal(HtmlConversionProfile.PositionedReview,
            PowerPointHtmlSaveOptions.CreateVisualReviewProfile().SharedProfile);
        Assert.Equal(HtmlConversionProfile.Document,
            WordToHtmlOptions.CreateDocumentRoundTripProfile().SharedProfile);
        Assert.Equal(HtmlConversionProfile.HighFidelityPrint,
            RtfToHtmlOptions.CreatePrintReviewProfile().SharedProfile);

        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new ExcelHtmlSaveOptions().Profile = OfficeHtmlConversionProfile.WordPrintReview);
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new PowerPointHtmlSaveOptions().Profile = OfficeHtmlConversionProfile.ExcelVisualReview);
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new WordToHtmlOptions().Profile = OfficeHtmlConversionProfile.PowerPointVisualReview);
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new RtfToHtmlOptions().Profile = OfficeHtmlConversionProfile.WordSemanticDocument);
    }

    [Fact]
    public void TargetCapabilityContracts_KeepDirectionalProfilesSeparate() {
        HtmlTargetCapabilityContract word = HtmlTargetCapabilityContracts.Get(HtmlConversionTarget.Word);
        HtmlTargetCapabilityContract pdf = HtmlTargetCapabilityContracts.Get(HtmlConversionTarget.Pdf);
        HtmlTargetCapabilityContract image = HtmlTargetCapabilityContracts.Get(HtmlConversionTarget.Image);

        Assert.Equal(new[] { "OfficeIMO", "UntrustedHtml", "TrustedDocument" }, word.HtmlToTarget.Profiles);
        Assert.Equal(new[] { "SemanticDocument", "DocumentRoundTrip", "PrintReview" }, word.TargetToHtml!.Profiles);
        Assert.Equal(new[] { "PagedPrint" }, pdf.HtmlToTarget.Profiles);
        Assert.Equal(new[] { "Semantic", "PositionedReview" }, pdf.TargetToHtml!.Profiles);
        Assert.Null(image.TargetToHtml);
        Assert.DoesNotContain("PositionedReview", pdf.HtmlToTarget.Profiles);
        Assert.DoesNotContain("PagedPrint", pdf.TargetToHtml.Profiles);
        HtmlToTargetCapabilityContract pdfImport = pdf.HtmlToTarget;
        TargetToHtmlCapabilityContract pdfExport = pdf.TargetToHtml;
        Assert.Equal(HtmlCapabilitySupportLevel.Approximated, pdfImport.GetSupport(HtmlSemanticFeature.Forms));
        Assert.Equal(HtmlCapabilitySupportLevel.Supported, pdfExport.GetSupport(HtmlSemanticFeature.Forms));
        Assert.NotEqual(pdfImport.EntryPoint, pdfExport.EntryPoint);
        Assert.NotEqual(pdfImport.ResultContract, pdfExport.ResultContract);
        Assert.NotEqual(pdfImport.IoAndAsyncBoundary, pdfExport.IoAndAsyncBoundary);
        Assert.NotEqual(pdfImport.DiagnosticsContract, pdfExport.DiagnosticsContract);
    }

    [Fact]
    public void DirectionalCapabilityAndCatalogCollections_AreDefensiveSnapshots() {
        string[] profiles = { "One" };
        HtmlSemanticFeature[] supported = Enum.GetValues(typeof(HtmlSemanticFeature)).Cast<HtmlSemanticFeature>().ToArray();
        var route = new HtmlToTargetCapabilityContract(
            "Convert", "Result", "Sync", "Structured diagnostics", profiles,
            supported, Array.Empty<HtmlSemanticFeature>(), Array.Empty<HtmlSemanticFeature>());

        profiles[0] = "Mutated";
        supported[0] = supported[1];

        Assert.Equal("One", route.Profiles[0]);
        Assert.Equal(HtmlCapabilitySupportLevel.Supported, route.GetSupport(HtmlSemanticFeature.Metadata));
        Assert.Throws<NotSupportedException>(() => ((IList<string>)route.Profiles)[0] = "Changed");
        Assert.Throws<NotSupportedException>(() =>
            ((IList<HtmlTargetCapabilityContract>)HtmlTargetCapabilityContracts.All).RemoveAt(0));

        var artifactSource = new List<HtmlCapabilityGalleryArtifact> {
            new HtmlCapabilityGalleryArtifact("source", "html", "source.html", "text/html", 1, new string('0', 64))
        };
        var diagnosticSource = new List<HtmlDiagnostic> {
            new HtmlDiagnostic("OfficeIMO.Tests", "Snapshot", "snapshot")
        };
        var gallery = new HtmlCapabilityGalleryResult(
            new HtmlCapabilityGalleryScenario("snapshot", "Snapshot", "HTML", "Snapshot proof"),
            artifactSource,
            diagnosticSource);
        artifactSource.Clear();
        diagnosticSource.Clear();

        Assert.Single(gallery.Artifacts);
        Assert.Single(gallery.Diagnostics);
        Assert.True(gallery.IsReadOnly);
        Assert.True(gallery.Diagnostics.IsReadOnly);
        Assert.Throws<NotSupportedException>(() => ((IList<HtmlCapabilityGalleryArtifact>)gallery.Artifacts).Clear());
        Assert.Throws<InvalidOperationException>(() => gallery.AddArtifact(
            new HtmlCapabilityGalleryArtifact("late", "html", "late.html", "text/html", 1, new string('1', 64))));
        Assert.Throws<InvalidOperationException>(() => gallery.Diagnostics.Clear());
        Assert.Throws<InvalidOperationException>(() => gallery.Diagnostics.Add(
            new HtmlDiagnostic("OfficeIMO.Tests", "Late", "late")));

        var builder = new HtmlCapabilityGalleryResult(
            new HtmlCapabilityGalleryScenario("builder", "Builder", "HTML", "Compatibility builder"));
        builder.AddArtifact(new HtmlCapabilityGalleryArtifact(
            "initial", "html", "initial.html", "text/html", 1, new string('2', 64)));
        builder.Diagnostics.Add(new HtmlDiagnostic("OfficeIMO.Tests", "Initial", "initial"));
        var manifest = new HtmlCapabilityGalleryManifest(
            builder,
            HtmlConversionProfile.Document,
            roundTripScore: null,
            resourceManifest: null);

        builder.AddArtifact(new HtmlCapabilityGalleryArtifact(
            "later", "html", "later.html", "text/html", 1, new string('3', 64)));
        builder.Diagnostics.Add(new HtmlDiagnostic("OfficeIMO.Tests", "Later", "later"));

        Assert.Single(manifest.Result.Artifacts);
        Assert.Single(manifest.Result.Diagnostics);
        Assert.True(manifest.Result.IsReadOnly);
        Assert.True(manifest.Result.Diagnostics.IsReadOnly);
    }

    [Fact]
    public void DocumentOutputSettings_ComposeShellPolicyAndKeepCompatibilityAliasesSynchronized() {
        var excel = ExcelHtmlSaveOptions.CreateVisualReviewProfile();
        excel.DocumentOutput.EmitDocumentShell = false;
        excel.DocumentOutput.Language = "pl-PL";
        excel.DocumentOutput.NewLine = "\r\n";
        excel.Title = "Review";

        Assert.False(excel.EmitDocumentShell);
        Assert.Equal("pl-PL", excel.Language);
        Assert.Equal("\r\n", excel.NewLine);
        Assert.Equal("Review", excel.DocumentOutput.Title);
        Assert.True(excel.IncludeDefaultStyles);
        Assert.Equal(OfficeVisualThemeKind.Report, excel.Theme);

        var rtf = RtfToHtmlOptions.CreateWebSafeProfile();
        Assert.True(rtf.FragmentOnly);
        rtf.FragmentOnly = false;
        Assert.True(rtf.DocumentOutput.EmitDocumentShell);

        OfficeHtmlDocumentOptions clone = excel.DocumentOutput.Clone();
        clone.Title = "Clone";
        Assert.Equal("Review", excel.Title);
    }

    [Fact]
    public void DocumentOutputSettings_DriveRealWordAndExcelDocumentAndFragmentOutput() {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph("Fragment body");
        WordToHtmlOptions wordOptions = WordToHtmlOptions.CreateSemanticDocumentProfile();
        wordOptions.EmitDocumentShell = false;
        wordOptions.AdditionalMetaTags.Add(("review-id", "discarded-in-fragment"));

        HtmlTextConversionResult wordResult = word.ToHtmlResult(wordOptions);

        Assert.StartsWith("<style", wordResult.RequireValue(), StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Fragment body", wordResult.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("<html", wordResult.Value, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("review-id", wordResult.Value, StringComparison.Ordinal);
        Assert.Contains(wordResult.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "DocumentHeadMetadataOmittedForFragment");

        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        workbook.AddWorksheet("Data").CellValue(1, 1, "Value");
        ExcelHtmlSaveOptions excelOptions = ExcelHtmlSaveOptions.CreateSemanticTablesProfile();
        excelOptions.Title = "Packed review";
        excelOptions.Language = "pl-PL";
        excelOptions.NewLine = "\r\n";

        string documentHtml = workbook.ToHtml(excelOptions);
        Assert.StartsWith("<!doctype html>\r\n<html lang=\"pl-PL\">", documentHtml, StringComparison.Ordinal);
        Assert.Contains("<title>Packed review</title>", documentHtml, StringComparison.Ordinal);

        excelOptions.EmitDocumentShell = false;
        string fragmentHtml = workbook.ToHtml(excelOptions);
        Assert.DoesNotContain("<!doctype", fragmentHtml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<html", fragmentHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<table", fragmentHtml, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DocumentOutputSettings_PreserveRequiredAndCallerBodyClassesAcrossAdapters() {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph("Word");
        WordToHtmlOptions wordOptions = WordToHtmlOptions.CreateSemanticDocumentProfile();
        wordOptions.DocumentOutput.BodyClass = "customer-shell officeimo-html customer-shell";
        Assert.Contains("class=\"officeimo-html officeimo-word-html customer-shell\"", word.ToHtml(wordOptions), StringComparison.Ordinal);

        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        workbook.AddWorksheet("Data").CellValue(1, 1, "Excel");
        ExcelHtmlSaveOptions excelOptions = ExcelHtmlSaveOptions.CreateSemanticTablesProfile();
        excelOptions.DocumentOutput.BodyClass = "customer-shell";
        Assert.Contains("<body class=\"officeimo-html officeimo-excel-html customer-shell\"", workbook.ToHtml(excelOptions), StringComparison.Ordinal);

        string presentationPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
        try {
            using PowerPointPresentation presentation = PowerPointPresentation.Create(presentationPath);
            presentation.AddSlide().AddTextBox("PowerPoint");
            PowerPointHtmlSaveOptions powerPointOptions = PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile();
            powerPointOptions.DocumentOutput.BodyClass = "customer-shell";
            Assert.Contains("<body class=\"officeimo-html officeimo-powerpoint-html customer-shell\"", presentation.ToHtml(powerPointOptions), StringComparison.Ordinal);
        } finally {
            if (File.Exists(presentationPath)) File.Delete(presentationPath);
        }

        RtfDocument rtf = RtfDocument.Create();
        rtf.AddParagraph("RTF");
        RtfToHtmlOptions rtfOptions = RtfToHtmlOptions.CreatePrintReviewProfile();
        rtfOptions.DocumentOutput.BodyClass = "customer-shell";
        Assert.Contains("<body class=\"officeimo-html officeimo-rtf-html customer-shell\"", rtf.ToHtml(rtfOptions), StringComparison.Ordinal);

    }

    [Fact]
    public void WordFieldReviewPolicy_PreservesStoredResultAndMakesInstructionsInertOptInMetadata() {
        using WordDocument document = WordDocument.Create();
        document.BuiltinDocumentProperties.Creator = "Stored Author";
        document.AddParagraph("Author: ").AddField(WordFieldType.Author);

        HtmlTextConversionResult visible = document.ToHtmlResult(new WordToHtmlOptions {
            FieldPolicy = WordFieldExportPolicy.VisibleResult
        });
        HtmlTextConversionResult review = document.ToHtmlResult(new WordToHtmlOptions {
            FieldPolicy = WordFieldExportPolicy.VisibleResultWithReviewMetadata
        });

        Assert.DoesNotContain("data-officeimo-field-instruction", visible.Value, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-field-instruction", review.Value, StringComparison.Ordinal);
        Assert.Contains("AUTHOR", review.Value, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("never evaluates Word fields", review.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("<script", review.Value, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WordHtmlReview_ExportsFloatingPictureGeometryAndEffectsAsDiagnosedInertMetadata() {
        using WordDocument document = WordDocument.Create();
        byte[] pixels = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABFSURBVEhLY1BNfv2flpgBXYDaeBhaILCzkSKMbt6oBRgY3bxRCzAwunmjFmBgdPNGLcDA6OaNWoCB0c3DsIDaeNQCghgAFxBXzP1LTe4AAAAASUVORK5CYII=");
        using var imageStream = new MemoryStream(pixels);
        WordParagraph paragraph = document.AddParagraph();
        paragraph.AddImage(imageStream, "floating.png", 48, 36, WordImageTextWrapping.Square, "Floating marker");
        WordImage image = Assert.IsType<WordImage>(paragraph.Image);
        image.HorizontalPositionRelativeFrom = WordHorizontalRelativePosition.Page;
        image.VerticalPositionRelativeFrom = WordVerticalRelativePosition.Page;
        image.HorizontalPositionOffset = 914400;
        image.VerticalPositionOffset = 457200;
        image.Rotation = 15;
        image.LuminanceBrightness = 20;

        HtmlTextConversionResult result = document.ToHtmlResult(WordToHtmlOptions.CreatePrintReviewProfile());

        Assert.Contains("data-officeimo-anchor=\"floating\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-wrap=\"Square\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-horizontal-offset-emu=\"914400\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-vertical-offset-emu=\"457200\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-rotation=\"15\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-brightness=\"20\"", result.Value, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "FloatingImageProjectedForReview" &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }

    [Fact]
    public void WordHtmlReview_ProjectsTrackedChangesAcrossStoriesWithoutMutatingSource() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Body ").AddInsertedText("new", "Reviewer").AddDeletedText("old", "Reviewer");
        document.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default)
            .AddParagraph("Header ").AddInsertedText("new", "Reviewer").AddDeletedText("old", "Reviewer");
        string bodyXml = document._document.OuterXml;
        string headerXml = document.Sections[0].Header.Default!._header!.OuterXml;

        string finalHtml = document.ToHtml(new WordToHtmlOptions {
            ExportHeadersAndFooters = true,
            TrackedChangePolicy = WordTrackedChangeExportPolicy.Final
        });
        string originalHtml = document.ToHtml(new WordToHtmlOptions {
            ExportHeadersAndFooters = true,
            TrackedChangePolicy = WordTrackedChangeExportPolicy.Original
        });
        string markupHtml = document.ToHtml(WordToHtmlOptions.CreateDocumentRoundTripProfile());

        Assert.Contains("Body new", finalHtml, StringComparison.Ordinal);
        Assert.Contains("Header new", finalHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("Body old", finalHtml, StringComparison.Ordinal);
        Assert.Contains("Body old", originalHtml, StringComparison.Ordinal);
        Assert.Contains("Header old", originalHtml, StringComparison.Ordinal);
        Assert.DoesNotContain("Body new", originalHtml, StringComparison.Ordinal);
        Assert.Contains("<ins", markupHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<del", markupHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data-officeimo-review-policy=\"markup\"", markupHtml, StringComparison.Ordinal);
        Assert.Equal(bodyXml, document._document.OuterXml);
        Assert.Equal(headerXml, document.Sections[0].Header.Default!._header!.OuterXml);
    }

    [Fact]
    public void WordHtmlReview_ScopesInventoriesToEnabledStoriesAndSelectedTrackedView() {
        using WordDocument document = WordDocument.Create();
        WordParagraph body = document.AddParagraph("Body field: ");
        body.AddField(WordFieldType.Author);
        body.AddInsertedText("body revision", "Reviewer");
        WordParagraph header = document.Sections[0]
            .GetOrCreateHeader(WordHeaderFooterType.Default)
            .AddParagraph("PRIVATE HEADER ");
        header._paragraph.Append(new SimpleField(new Run(new Text("PRIVATE HEADER RESULT"))) {
            Instruction = "PRIVATE_HEADER_FIELD"
        });
        header.AddInsertedText("PRIVATE HEADER REVISION", "Reviewer");

        var deletedField = new DeletedRun {
            Author = "Reviewer",
            Date = DateTime.UtcNow,
            Id = "9001"
        };
        deletedField.Append(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" PRIVATE_DELETED_FIELD ")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new DeletedText("PRIVATE DELETED RESULT")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        body._paragraph.Append(deletedField);

        HtmlTextConversionResult result = document.ToHtmlResult(new WordToHtmlOptions {
            ExportHeadersAndFooters = false,
            TrackedChangePolicy = WordTrackedChangeExportPolicy.Markup,
            FieldPolicy = WordFieldExportPolicy.VisibleResultWithReviewMetadata
        });
        HtmlTextConversionResult finalResult = document.ToHtmlResult(new WordToHtmlOptions {
            ExportHeadersAndFooters = false,
            TrackedChangePolicy = WordTrackedChangeExportPolicy.Final,
            FieldPolicy = WordFieldExportPolicy.VisibleResultWithReviewMetadata
        });

        Assert.Contains("body revision", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-field-location=\"Body\"", result.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("PRIVATE HEADER", result.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("PRIVATE_DELETED_FIELD", finalResult.Value, StringComparison.Ordinal);
        Assert.DoesNotContain("PRIVATE DELETED RESULT", finalResult.Value, StringComparison.Ordinal);
    }

    [Fact]
    public void WordHtmlReview_DetectsFieldsInEnabledFootnotesAndEndnotes() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Footnote").AddFootNote("footnote value");
        document.AddParagraph("Endnote").AddEndNote("endnote value");
        Assert.NotNull(document.FootNotes[0].Paragraphs);
        Assert.NotNull(document.EndNotes[0].Paragraphs);
        document.FootNotes[0].Paragraphs!.Last().AddField(WordFieldType.Page);
        document.EndNotes[0].Paragraphs!.Last().AddField(WordFieldType.NumPages);

        HtmlTextConversionResult result = document.ToHtmlResult(new WordToHtmlOptions {
            ExportFootnotes = true,
            ExportEndnotes = true,
            FieldPolicy = WordFieldExportPolicy.VisibleResultWithReviewMetadata
        });

        Assert.Contains("data-officeimo-field-location=\"Footnote\"", result.Value, StringComparison.Ordinal);
        Assert.Contains("data-officeimo-field-location=\"Endnote\"", result.Value, StringComparison.Ordinal);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "FieldInstructionsFlattened");
    }

    [Fact]
    public void WordHtmlReview_BudgetsFieldAndRevisionInventoriesAndRestoresSourceOnFailure() {
        using WordDocument fieldDocument = WordDocument.Create();
        fieldDocument.AddParagraph()._paragraph.Append(new SimpleField(new Run(new Text("value"))) {
            Instruction = new string('F', 12_000)
        });
        HtmlConversionLimitException fieldException = Assert.Throws<HtmlConversionLimitException>(() =>
            fieldDocument.ToHtmlResult(new WordToHtmlOptions {
                FieldPolicy = WordFieldExportPolicy.VisibleResultWithReviewMetadata,
                MaxOutputCharacters = 8_000
            }));
        Assert.Equal("FieldInventory:instruction", fieldException.LimitSource);

        using WordDocument revisionDocument = WordDocument.Create();
        revisionDocument.AddParagraph().AddInsertedText(new string('R', 6_000), "Reviewer");
        string sourceXml = revisionDocument._document.OuterXml;
        HtmlConversionLimitException revisionException = Assert.Throws<HtmlConversionLimitException>(() =>
            revisionDocument.ToHtmlResult(new WordToHtmlOptions {
                TrackedChangePolicy = WordTrackedChangeExportPolicy.Markup,
                MaxOutputCharacters = 9_000
            }));
        Assert.Equal("RevisionInventory:text", revisionException.LimitSource);
        Assert.Equal(sourceXml, revisionDocument._document.OuterXml);
    }

    [Fact]
    public void PdfHtmlProfileContracts_ExposeDefensiveReadOnlySnapshots() {
        PdfHtmlProfileContract semantic = PdfHtmlProfileContracts.Get(PdfHtmlProfile.Semantic);
        var values = Assert.IsAssignableFrom<IList<string>>(semantic.PreservedSignals);

        Assert.Throws<NotSupportedException>(() => values[0] = "mutated");
        Assert.Equal("metadata", PdfHtmlProfileContracts.Get(PdfHtmlProfile.Semantic).PreservedSignals[0]);
    }

    [Fact]
    public void ExcelHtmlReview_ExposesPivotDefinitionWithoutExecutingInteractiveBehavior() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
        try {
            using ExcelDocument document = ExcelDocument.Create(path);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Region");
            sheet.CellValue(1, 2, "Sales");
            sheet.CellValue(2, 1, "East");
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(3, 1, "West");
            sheet.CellValue(3, 2, 12);
            sheet.AddPivotTable("A1:B3", "D2", "SalesPivot",
                rowFields: new[] { "Region" },
                dataFields: new[] { new ExcelPivotDataField("Sales", ExcelPivotDataFunction.Sum, "Total Sales") });

            HtmlTextConversionResult result = document.ToHtmlResult(ExcelHtmlSaveOptions.CreateSemanticTablesProfile());
            string html = result.Value;

            Assert.Contains("data-officeimo-feature=\"pivot-table\"", html, StringComparison.Ordinal);
            Assert.Contains("SalesPivot", html, StringComparison.Ordinal);
            Assert.Contains("Total Sales", html, StringComparison.Ordinal);
            Assert.Contains("refresh, drill, caches, slicers, and timelines remain native workbook behavior", html, StringComparison.Ordinal);
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == HtmlConversionDiagnosticCodes.ExcelPivotReviewApproximated &&
                diagnostic.LossKind == OfficeConversionLossKind.Approximation);
            Assert.True(result.HasLoss);
            Assert.Throws<HtmlConversionException>(() => result.RequireNoLoss());
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void PowerPointHtmlReview_ExposesMastersSmartArtMediaPosterAndStaticPolicy() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".pptx");
        try {
            using PowerPointPresentation presentation = PowerPointPresentation.Create(path);
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddSmartArt(PowerPointSmartArtType.BasicProcess, new[] { "Discover", "Convert", "Review" });
            PowerPointMedia media = slide.AddLinkedVideo(new Uri("https://example.test/review-video.mp4"));
            media.LuminanceBrightness = 15;

            PowerPointToHtmlResult result = presentation.ToHtmlResult(PowerPointHtmlSaveOptions.CreateSemanticSlidesProfile());
            string html = result.Value;

            Assert.Contains("data-officeimo-feature=\"slide-master\"", html, StringComparison.Ordinal);
            Assert.Contains("data-officeimo-feature=\"smartart\"", html, StringComparison.Ordinal);
            Assert.Contains("Discover", html, StringComparison.Ordinal);
            Assert.Contains("data-officeimo-feature=\"media\"", html, StringComparison.Ordinal);
            Assert.Contains("data-officeimo-poster-frame=\"true\"", html, StringComparison.Ordinal);
            Assert.Contains("Audio and video are never executed", html, StringComparison.Ordinal);
            Assert.DoesNotContain("<script", html, StringComparison.OrdinalIgnoreCase);
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == HtmlConversionDiagnosticCodes.PowerPointMasterReviewApproximated);
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == HtmlConversionDiagnosticCodes.PowerPointSmartArtReviewApproximated);
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == HtmlConversionDiagnosticCodes.PowerPointMediaReviewApproximated);
            Assert.Contains(result.Report.Diagnostics, diagnostic =>
                diagnostic.Code == HtmlConversionDiagnosticCodes.PowerPointEffectReviewApproximated);
            Assert.True(result.HasLoss);
            Assert.Throws<HtmlConversionException>(() => result.RequireNoLoss());
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
