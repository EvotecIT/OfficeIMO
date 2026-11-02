using OfficeIMO.IWork;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Pages_fully_overridden_fractional_table_defaults_do_not_require_partial_DOCX(bool allowPartial) {
        using var package = EffectivePagesDimensions();
        using var result = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new IWorkConversionOptions {
                Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = allowPartial
            });

        Assert.False(result.IsVisualFallback);
        Assert.False(result.Report.IsPartialEditableReconstruction);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_PAGES_DOCX_PRECISION");
        result.Report.RequireCompleteEditableReconstruction();
        Assert.Equal(10.125, Assert.Single(result.Projection.Tables).DefaultRowHeight);
        Assert.Equal(20.125, Assert.Single(result.Projection.Tables).DefaultColumnWidth);
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        var table = Assert.Single(reopened.Tables);
        Assert.Equal(new[] { 405, 200, 600 }, table.RowHeight.Select(height => checked((int)(height?.ToTwips() ?? 0))));
        Assert.Equal(new[] { 810, 400 }.Select(value => (WordTableWidth?)WordTableWidth.Length(OfficeLength.Emus((long)value * 635))), table.ColumnWidth);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Pages_active_fractional_table_default_still_requires_partial_DOCX(bool missingRow) {
        using var package = EffectivePagesDimensions(missingRow, missingColumn: !missingRow);
        var source = IWorkSourceDocument.Open(package);
        Assert.Contains("sizing", Assert.Throws<InvalidDataException>(() => source.ToWordDocumentResult(
            new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly })).Message);

        using var partial = source.ToWordDocumentResult(new IWorkConversionOptions {
            Mode = IWorkConversionMode.EditableOnly, AllowPartialEditableReconstruction = true
        });
        Assert.True(partial.Report.IsPartialEditableReconstruction);
        Assert.Contains(partial.Report.Diagnostics, d => d.Code == "IWORK_PAGES_DOCX_PRECISION");
        using var saved = new MemoryStream(); partial.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        var table = Assert.Single(reopened.Tables);
        Assert.Equal(missingRow ? new[] { 405, 203, 600 } : new[] { 405, 200, 600 }, table.RowHeight.Select(height => checked((int)(height?.ToTwips() ?? 0))));
        Assert.Equal((missingRow ? new[] { 810, 400 } : new[] { 810, 403 }).Select(value => (WordTableWidth?)WordTableWidth.Length(OfficeLength.Emus((long)value * 635))), table.ColumnWidth);
    }

    private static MemoryStream EffectivePagesDimensions(bool missingRow = false, bool missingColumn = false) =>
        DimensionPackage(IWorkDocumentKind.Pages,
            firstBucketPayload: Message(VarintField(1, 1), BytesField(2, DimensionHeader(0, 20.25f)),
                missingRow ? Array.Empty<byte>() : BytesField(2, DimensionHeader(1, 10))),
            columnBucketPayload: Message(VarintField(1, 1), BytesField(2, DimensionHeader(0, 40.5f)),
                missingColumn ? Array.Empty<byte>() : BytesField(2, DimensionHeader(1, 20))),
            defaultRowHeight: 10.125, defaultColumnWidth: 20.125);
}
