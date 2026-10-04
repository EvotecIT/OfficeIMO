using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Source_units_identify_colliding_tables_without_retaining_payloads() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Same", 1, 1, 42d), new TableSpec("Same", 1, 1, 84d)
        });
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            readOptions: new IWorkReadOptions { PreserveSourceRecords = false },
            conversionOptions: new IWorkConversionOptions { NormalizeWorksheetNames = true });
        Assert.Empty(result.Report.PreservedRecords);
        IWorkSourceUnitCount tables = Count(result.Report, IWorkSourceUnitKind.Table);
        Assert.Equal(2, tables.TotalCount);
        Assert.Equal(2, tables.ReconstructedCount);
        Assert.Equal(0, tables.OmittedCount);
        IWorkTable[] projected = result.Projection.Sheets.SelectMany(sheet => sheet.Tables).ToArray();
        Assert.NotEqual(projected[0].SourceIdentity!.RecordIdentifier, projected[1].SourceIdentity!.RecordIdentifier);
        foreach (IWorkTable table in projected) {
            IWorkSourceUnit unit = Assert.Single(result.Report.SourceUnits, unit => unit.Identity.RecordIdentifier == table.SourceIdentity!.RecordIdentifier);
            Assert.Equal(IWorkSourceUnitDisposition.Reconstructed, unit.Disposition);
            IWorkArchiveRecord record = Assert.Single(result.Source.Records, record => record.IsPrimary && record.Identifier == unit.Identity.RecordIdentifier);
            Assert.Equal(record.EntryPath, unit.Identity.EntryPath);
            Assert.Equal(record.MessageType, unit.Identity.MessageType);
        }
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        Assert.Equal(tables.ReconstructedCount, reopened.Sheets.Count);
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
        Assert.Equal(84d, reopened.Sheets[1].CellAt(1, 1).GetValue<double>());
    }

    [Fact]
    public void Partial_table_reconstruction_reports_the_selected_omitted_table() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Recovered", 1, 1, 42d), new TableSpec("Missing", 1, 1, 0d, missingModel: true)
        }, includePreview: true);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.False(result.IsVisualFallback);
        IWorkSourceUnitCount tables = Count(result.Report, IWorkSourceUnitKind.Table);
        Assert.Equal(2, tables.TotalCount);
        Assert.Equal(1, tables.ReconstructedCount);
        Assert.Equal(1, tables.OmittedCount);
        IWorkSourceUnit omitted = Assert.Single(result.Report.SourceUnits, unit => unit.Disposition == IWorkSourceUnitDisposition.Omitted);
        Assert.Contains(result.Report.Diagnostics, diagnostic => diagnostic.RecordIdentifier == omitted.Identity.RecordIdentifier);
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        Assert.Single(reopened.Sheets);
        Assert.Equal(42d, reopened.Sheets[0].CellAt(1, 1).GetValue<double>());
    }

    [Fact]
    public void Independent_keynote_slides_retain_their_native_identities() {
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(CorpusFixture("nim-iwork/simple.key"),
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true });
        Assert.False(result.IsVisualFallback);
        Assert.Equal(2, Count(result.Report, IWorkSourceUnitKind.Slide).ReconstructedCount);
        Assert.Equal(2, result.Projection.Slides.Select(slide => slide.SourceIdentity!.RecordIdentifier).Distinct().Count());
        Assert.All(result.Projection.Slides, slide => Assert.Contains(result.Report.SourceUnits,
            unit => unit.Kind == IWorkSourceUnitKind.Slide && unit.Identity.RecordIdentifier == slide.SourceIdentity!.RecordIdentifier
                && unit.Disposition == IWorkSourceUnitDisposition.Reconstructed));
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using PowerPointPresentation reopened = PowerPointPresentation.Load(saved);
        Assert.Equal(2, reopened.Slides.Count);
    }

    [Fact]
    public void A_visual_preview_does_not_establish_individual_source_unit_coverage() {
        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(CorpusFixture("nim-iwork/simple.key"),
            conversionOptions: IWorkTestPolicy.ForIncompletePreview(new IWorkConversionOptions { Mode = IWorkConversionMode.VisualOnly }));
        Assert.True(result.IsVisualFallback);
        Assert.False(result.Report.HasCompleteVisualCoverage);
        Assert.NotEmpty(result.Report.SourceUnits);
        Assert.All(result.Report.SourceUnits, unit => Assert.Equal(IWorkSourceUnitDisposition.Unassessed, unit.Disposition));
        Assert.Equal(2, Count(result.Report, IWorkSourceUnitKind.Slide).UnassessedCount);
        Assert.Equal(0, result.Report.ReconstructedItemCount);
    }

    [Fact]
    public void Unused_header_records_are_excluded_from_document_unit_counts() {
        using MemoryStream package = CreatePagesPackageWithTwoSections(emptySecondSection: true);
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package);
        IWorkSourceUnitCount text = Count(result.Report, IWorkSourceUnitKind.Text);
        Assert.Equal(2, text.TotalCount);
        Assert.Equal(2, text.ReconstructedCount);
        Assert.Equal(0, text.OmittedCount);
        Assert.Equal(0, text.UnassessedCount);
        Assert.Contains(result.Source.Records, record => record.Identifier == 8 && record.IsPrimary);
        Assert.DoesNotContain(result.Report.SourceUnits, unit => unit.Identity.RecordIdentifier is 8 or 9);
        Assert.NotNull(result.Projection.Body.SourceIdentity);
        using var saved = new MemoryStream();
        result.Value.Save(saved);
        saved.Position = 0;
        using WordDocument reopened = WordDocument.Load(saved);
        Assert.Contains(reopened.Sections[0].Header.Default!.Paragraphs, paragraph => paragraph.Text == "First header");
        Assert.DoesNotContain(reopened.Sections[1].Header.Default!.Paragraphs, paragraph => paragraph.Text == "Second header");
    }

    private static IWorkSourceUnitCount Count(IWorkConversionReport report, IWorkSourceUnitKind kind) =>
        Assert.Single(report.SourceUnitCounts, count => count.Kind == kind);
}
