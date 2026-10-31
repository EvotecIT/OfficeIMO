using OfficeIMO.Core.Internal;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelLegacyVbaProducerPreservationTests {
    [Fact]
    public void Native_xls_cell_addition_retains_independent_chart_and_macro_carriers() {
        string path = Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.TestAssets", "Documents", "LegacyXlsCorpus", "openpreserve-format-corpus", "valid.xls");
        byte[] source = File.ReadAllBytes(path);
        Assert.True(OfficeCompoundFileReader.TryRead(source, out OfficeCompoundFile? original, out string? sourceError), sourceError);
        var originalVba = original!.Streams.Where(item => item.Key.StartsWith("_VBA_PROJECT_CUR/", StringComparison.OrdinalIgnoreCase)).ToDictionary(item => item.Key, item => item.Value);
        Assert.NotEmpty(originalVba);
        using var document = ExcelDocument.LoadLegacyXls(new MemoryStream(source),
            new LegacyXlsImportOptions { PreserveExternalWorkbookLinks = true });
        document.Sheets[0].CellValue(1, 1, "Cell edit with original VBA retained");
        Assert.NotEmpty(document.LegacyXlsChartSheets);
        using var destination = new MemoryStream();
        document.Save(destination, ExcelFileFormat.Xls);
        using var reloaded = ExcelDocument.Load(new MemoryStream(destination.ToArray()));
        Assert.Equal("Cell edit with original VBA retained", reloaded.Sheets[0].CellAt(1, 1).GetValue<string>());
        Assert.Equal(document.LegacyXlsChartSheets.Count, reloaded.LegacyXlsChartSheets.Count);
        Assert.True(OfficeCompoundFileReader.TryRead(destination.ToArray(), out OfficeCompoundFile? saved, out string? error), error);
        foreach (var stream in originalVba) Assert.Equal(stream.Value, saved!.Streams[stream.Key]);
    }
}
