using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_keynote_conversion_preserves_both_independent_source_slides_without_partial_policy() {
        string path = CorpusFixture("nim-iwork/simple.key");
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(path);

        Assert.False(result.IsVisualFallback, string.Join("; ", result.Report.Diagnostics.Select(diagnostic => diagnostic.ToString())));
        Assert.Equal(2, result.Value.Slides.Count);
        Assert.NotEmpty(result.Value.Slides.SelectMany(slide => slide.TextBoxes));
        Assert.False(result.Report.IsPartialEditableReconstruction);
        result.Report.RequireCompleteEditableReconstruction();
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_PARAGRAPH_PAGINATION_UNSUPPORTED");
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
        Assert.Equal(2, reopened.Slides.Count);
        Assert.NotEmpty(reopened.Slides.SelectMany(slide => slide.TextBoxes));
    }

    [Fact]
    public void Partial_pages_conversion_keeps_independent_text_tables_and_image() {
        using var result = WordIWorkConverter.ConvertPagesToWordResult(CorpusFixture("picodocs/sample-v14.4.pages"),
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });

        Assert.False(result.IsVisualFallback, string.Join("; ", result.Report.Diagnostics.Select(diagnostic => diagnostic.ToString())));
        Assert.Equal(3, result.Value.Tables.Count);
        Assert.Single(result.Value.Images);
        Assert.Equal(3, Assert.Single(result.Report.SourceUnitCounts, count => count.Kind == IWorkSourceUnitKind.Table).ReconstructedCount);
        Assert.Equal(1, Assert.Single(result.Report.SourceUnitCounts, count => count.Kind == IWorkSourceUnitKind.Image).ReconstructedCount);
        IWorkObjectIdentity imageIdentity = Assert.Single(result.Projection.Images).SourceIdentity!;
        Assert.Contains(result.Report.SourceUnits, unit => unit.Identity.RecordIdentifier == imageIdentity.RecordIdentifier
            && unit.Kind == IWorkSourceUnitKind.Image && unit.Disposition == IWorkSourceUnitDisposition.Reconstructed);
        Assert.True(result.Report.IsPartialEditableReconstruction);
        Assert.Contains(result.Value.Paragraphs, paragraph => paragraph.Text.Length > 0);
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Equal(3, reopened.Tables.Count);
        Assert.Single(reopened.Images);
        Assert.Contains(reopened.Tables.SelectMany(table => table.Rows).SelectMany(row => row.Cells)
            .SelectMany(cell => cell.Paragraphs), paragraph => paragraph.Text == "Feature");
    }

    [Fact]
    public void Inline_object_markers_are_distinguished_from_invalid_source_text() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null,
            includePreview: true, bodyText: "Before \ufffc After");
        IWorkPagesProjection pages = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages).ReadPages();

        Assert.True(pages.Body.HasUnresolvedInlineObjects);
        Assert.False(pages.Body.HasInvalidSourceText);
        Assert.True(pages.Body.IsFormattingComplete);
        Assert.Contains(pages.Diagnostics, diagnostic => diagnostic.Message.Contains("unresolved inline objects"));
        Assert.DoesNotContain(pages.Diagnostics, diagnostic => diagnostic.Message.Contains("invalid UTF-8"));
    }

    [Theory]
    [InlineData("nim-iwork/simple.pages", IWorkDocumentKind.Pages)]
    [InlineData("nim-iwork/simple.numbers", IWorkDocumentKind.Numbers)]
    [InlineData("nim-iwork/simple.key", IWorkDocumentKind.Keynote)]
    public void Complete_visual_coverage_policy_rejects_first_page_or_composite_previews(string fixture,
        IWorkDocumentKind kind) {
        var options = new IWorkConversionOptions {
            Mode = IWorkConversionMode.VisualOnly, RequireCompleteVisualCoverage = true
        };
        Action convert = kind switch {
            IWorkDocumentKind.Pages => () => { using var result = WordIWorkConverter.ConvertPagesToWordResult(CorpusFixture(fixture), conversionOptions: options); },
            IWorkDocumentKind.Numbers => () => { using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(CorpusFixture(fixture), conversionOptions: options); },
            _ => () => { using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(CorpusFixture(fixture), conversionOptions: options); }
        };

        InvalidDataException error = Assert.Throws<InvalidDataException>(convert);
        Assert.Contains("complete document", error.Message);
    }

    [Fact]
    public void Editable_geometry_quantization_is_an_approximation_not_an_omission() {
        using MemoryStream package = CreateKeynotePackageWithRepeatedSlides(1,
            slideWidth: 960.0001f, slideHeight: 540f);
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.False(result.IsVisualFallback, string.Join("; ", result.Report.Diagnostics.Select(diagnostic => diagnostic.ToString())));
        IWorkDiagnostic precision = Assert.Single(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_PPTX_PRECISION");
        Assert.Equal(global::OfficeIMO.OfficeConversionLossKind.Approximation, precision.LossKind);
        Assert.Contains(result.Report.FidelityDiagnostics, diagnostic =>
            diagnostic.Code == precision.Code && diagnostic.LossKind == precision.LossKind);
    }

    [Fact]
    public void Numbers_name_mapping_preserves_colliding_tables_and_local_formulas_after_reopen() {
        string longName = new string('A', 40);
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec(longName + ":first", 1, 1, 42d, hasFormula: true, completeFormula: true),
            new TableSpec(longName + "/second", 1, 1, 84d)
        }, includePreview: true);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });

        Assert.False(result.IsVisualFallback);
        Assert.Equal(2, result.Value.Sheets.Count);
        Assert.NotEqual(result.Value.Sheets[0].Name, result.Value.Sheets[1].Name);
        Assert.All(result.Value.Sheets, sheet => Assert.InRange(sheet.Name.Length, 1, 31));
        Assert.All(result.WorksheetMappings, mapping => Assert.True(mapping.WasRenamed));
        Assert.Equal(1, result.WorksheetMappings[0].SourceTableIndex);
        Assert.Equal(2, result.WorksheetMappings[1].SourceTableIndex);
        Assert.Equal(longName + ":first", result.WorksheetMappings[0].SourceTableName);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NUMBERS_WORKSHEET_RENAMED"
            && diagnostic.LossKind == global::OfficeIMO.OfficeConversionLossKind.Approximation);
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using global::OfficeIMO.Excel.ExcelDocument reopened = global::OfficeIMO.Excel.ExcelDocument.Load(saved);
        Assert.Equal(result.WorksheetMappings.Select(mapping => mapping.DestinationName), reopened.Sheets.Select(sheet => sheet.Name));
        Assert.Equal(result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.Formula!.TrimStart('='), reopened.Sheets[0].GetFormulaText(1, 1));
        Assert.Equal(84d, reopened.Sheets[1].CellAt(1, 1).GetValue<double>());
    }

    [Fact]
    public void Partial_editable_policy_does_not_bypass_source_materialization_limits() {
        var conversion = new IWorkConversionOptions {
            Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true
        };
        var limits = new IWorkReadOptions { MaximumProjectedTables = 1 };
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("One", 1, 1, 1d), new TableSpec("Two", 1, 1, 2d)
        });

        InvalidDataException error = Assert.Throws<InvalidDataException>(() =>
            ExcelIWorkConverter.ConvertNumbersToExcelResult(package, limits, conversion));
        Assert.Contains("table count", error.Message);
    }

    [Fact]
    public void Invalid_source_text_remains_distinct_from_unresolved_inline_objects() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null,
            includePreview: true, bodyBytes: new byte[] { 0xc3, 0x28 });
        IWorkTextContent body = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages).ReadPages().Body;

        Assert.True(body.HasInvalidSourceText);
        Assert.False(body.HasUnresolvedInlineObjects);
        Assert.False(body.IsTextComplete);
    }

    [Fact]
    public void Partial_policy_still_rejects_oversized_word_destination_geometry() {
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null,
            includePreview: true, documentLayoutFields: PageLayoutFields(float.MaxValue));
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: IWorkTestPolicy.ForIncompletePreview(new IWorkConversionOptions { AllowPartialEditableReconstruction = true }));

        Assert.True(result.IsVisualFallback);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_PAGES_WORD_DESTINATION_UNSUPPORTED");
    }

    [Fact]
    public void Partial_policy_rejects_page_margins_that_round_to_an_empty_content_box() {
        byte[] layout = Message(FloatField(30, 612f), FloatField(31, 792f),
            FloatField(32, 305.99f), FloatField(33, 305.99f),
            FloatField(34, 72f), FloatField(35, 72f), FloatField(36, 36f), FloatField(37, 36f));
        using MemoryStream package = CreatePagesPackage(includeBody: true, textBox: null,
            includePreview: true, documentLayoutFields: layout);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: IWorkTestPolicy.ForIncompletePreview(new IWorkConversionOptions { AllowPartialEditableReconstruction = true }));

        Assert.True(result.IsVisualFallback);
    }

    [Theory]
    [InlineData(9)]
    [InlineData(10)]
    [InlineData(14)]
    public void Keynote_table_cell_pagination_is_reported_as_partial(int styleField) {
        using MemoryStream package = CreateFormulaTableWithRichCacheStyle(IWorkDocumentKind.Keynote,
            paginationStyleField: styleField);
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });

        Assert.False(result.IsVisualFallback);
        Assert.True(result.Report.IsPartialEditableReconstruction);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "IWORK_KEYNOTE_PARAGRAPH_PAGINATION_UNSUPPORTED");
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireCompleteEditableReconstruction());
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
        Assert.Contains("Styled", reopened.Slides[0].Tables.First().GetCell(0, 0).Text);
    }

    [Theory]
    [InlineData(22)]
    [InlineData(18)]
    public void Numbers_name_normalization_and_collision_suffix_preserve_valid_unicode(int prefixLength) {
        string name = new string('A', prefixLength) + "😀" + "tail";
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec(name, 1, 1, 1d), new TableSpec(name, 1, 1, 2d)
        });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using global::OfficeIMO.Excel.ExcelDocument reopened = global::OfficeIMO.Excel.ExcelDocument.Load(saved);

        Assert.Equal(2, reopened.Sheets.Count);
        Assert.NotEqual(reopened.Sheets[0].Name, reopened.Sheets[1].Name);
        Assert.All(reopened.Sheets, sheet => {
            System.Xml.XmlConvert.VerifyXmlChars(sheet.Name);
            Assert.InRange(sheet.Name.Length, 1, 31);
        });
        Assert.Equal(1d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(2d, reopened.Sheets[1].CellAt(1, 1).GetValue<double>());
    }

    private static string CorpusFixture(string path) => Path.Combine(AppContext.BaseDirectory,
        "Documents", "IWorkCorpus", path.Replace('/', Path.DirectorySeparatorChar));
}
