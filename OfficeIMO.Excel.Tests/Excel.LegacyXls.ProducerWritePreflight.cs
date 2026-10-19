using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelLegacyVbaProducerPreservationTests {
    [Fact]
    public void Native_xls_save_rejects_new_cells_in_an_independent_chart_macro_workbook_without_touching_output() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "LegacyXlsCorpus", "openpreserve-format-corpus", "valid.xls");
        byte[] source = File.ReadAllBytes(path);
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? original, out string? sourceError), sourceError);
        var originalVba = original!.Streams.Where(item => item.Key.StartsWith("_VBA_PROJECT_CUR/", StringComparison.OrdinalIgnoreCase)).ToDictionary(item => item.Key, item => item.Value);
        Assert.NotEmpty(originalVba);
        using var document = ExcelDocument.Load(new MemoryStream(source));
        document.Sheets[0].CellValue(1, 1, "Cell edit with original VBA retained");
        Assert.NotEmpty(document.LegacyXlsChartSheets);
        // A1 is absent in this source, so adding it exceeds the fixed-width literal preserving profile.
        using var destination = new MemoryStream(new byte[] { 1, 2, 3, 4 }, writable: true);
        var exception = Assert.Throws<NotSupportedException>(() => document.Save(destination, ExcelFileFormat.Xls,
            new ExcelSaveOptions { LossPolicy = OfficeConversionLossPolicy.Allow }));
        Assert.Contains("unsupported sheet types", exception.Message);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, destination.ToArray());
    }
}
