using System.IO.Compression;
using System.Xml.Linq;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.IWork.Internal;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Native_text_join_preserves_qualified_range_cache_storage_and_edited_recalculation() {
        using var manifest = ReadFunctionIdentityManifest();
        var expected = manifest.RootElement.GetProperty("textJoinCase");
        string sourceSheet = expected.GetProperty("sourceSheet").GetString()!;
        string sourceTable = expected.GetProperty("sourceTable").GetString()!;
        string targetSheet = expected.GetProperty("targetSheet").GetString()!;
        string targetTable = expected.GetProperty("targetTable").GetString()!;
        string address = expected.GetProperty("address").GetString()!;
        string cache = expected.GetProperty("cachedValue").GetString()!;
        Assert.Equal(expected.GetProperty("computedCurrentRangeValue").GetString(), cache);
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "numbers-parser", "cross-table-formulas.numbers");
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(path,
            conversionOptions: new IWorkConversionOptions { AllowPartialEditableReconstruction = true, NormalizeWorksheetNames = true });
        IWorkTableCell cell = result.Projection.Sheets.Single(s => s.Name == sourceSheet).Tables.Single(t => t.Name == sourceTable).GetCell(16, 1)!;
        Assert.True(cell.FormulaIsComplete);
        Assert.Equal("=TEXTJOIN(\",\",FALSE,'" + targetSheet.Replace("'", "''") + "'::'"
            + targetTable.Replace("'", "''") + "'::" + address + ")", cell.Formula);
        Assert.Equal(cache, cell.Value);
        Assert.DoesNotContain(result.Report.Diagnostics, d => d.Code == "IWORK_TABLE_FORMULA_PARTIAL");
        string destinationName = result.WorksheetMappings.Single(m => m.SourceSheetName == sourceSheet && m.SourceTableName == sourceTable).DestinationName;
        string targetName = result.WorksheetMappings.Single(m => m.SourceSheetName == targetSheet && m.SourceTableName == targetTable).DestinationName;
        string stored = "_xlfn.TEXTJOIN(\",\",FALSE,'" + targetName.Replace("'", "''") + "'!" + address + ")";
        using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using (var zip = new ZipArchive(output, ZipArchiveMode.Read, leaveOpen: true)) {
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            int index = result.Value.Sheets.ToList().FindIndex(s => s.Name == destinationName) + 1;
            using Stream xml = zip.GetEntry($"xl/worksheets/sheet{index}.xml")!.Open();
            XElement saved = XDocument.Load(xml).Descendants(ns + "c").Single(c => c.Attribute("r")?.Value == "A16");
            Assert.Equal(stored, saved.Element(ns + "f")!.Value);
            Assert.Equal("str", saved.Attribute("t")!.Value);
            Assert.Equal(cache, saved.Element(ns + "v")!.Value);
        }
        output.Position = 0;
        using var reopened = ExcelDocument.Load(output);
        ExcelSheet formulaSheet = reopened.Sheets.Single(s => s.Name == destinationName);
        Assert.Equal(stored, formulaSheet.GetFormulaText(16, 1));
        Assert.Equal(cache, formulaSheet.CellAt(16, 1).GetValue<string>());
        reopened.Sheets.Single(s => s.Name == targetName).CellValue(1, 2, "lion");
        Assert.True(reopened.Calculate() > 0);
        Assert.Equal("lion,dog,hamster", formulaSheet.CellAt(16, 1).GetValue<string>());
        reopened.Sheets.Single(s => s.Name == targetName).CellValue(1, 2, string.Empty);
        formulaSheet.CellFormula(18, 1, "TEXTJOIN(\",\",TRUE,'" + targetName.Replace("'", "''") + "'!" + address + ")");
        Assert.True(reopened.Calculate() > 0);
        Assert.Equal(",dog,hamster", formulaSheet.CellAt(16, 1).GetValue<string>());
        Assert.Equal("dog,hamster", formulaSheet.CellAt(18, 1).GetValue<string>());
    }

    [Fact]
    public void Text_join_respects_the_destination_text_argument_limit() {
        Assert.True(RenderFunction(328, 254).IsComplete);
        Assert.False(RenderFunction(328, 255).IsComplete);
    }

    [Fact]
    public void Text_join_prefix_expansion_requires_destination_fallback_without_truncation() {
        string payload = new string('x', 8190 - "TEXTJOIN(\"\",FALSE,\"\")".Length);
        byte[] formula = BytesField(1, Message(
            BytesField(1, Message(VarintField(1, 19), StringField(6, string.Empty))),
            BytesField(1, Message(VarintField(1, 18), VarintField(5, 0))),
            BytesField(1, Message(VarintField(1, 19), StringField(6, payload))),
            BytesField(1, Message(VarintField(1, 16), VarintField(2, 328), VarintField(3, 3)))));
        using var package = CreateNumbersPackage(new[] { new TableSpec("Functions", 1, 1, 42d,
            hasFormula: true, formulaPayload: formula) }, includePreview: true);
        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package, conversionOptions: new IWorkConversionOptions { RequireCompleteVisualCoverage = false });
        Assert.True(result.Projection.Sheets[0].Tables[0].GetCell(1, 1)!.FormulaIsComplete);
        Assert.True(result.IsVisualFallback);
        Assert.Empty(result.WorksheetMappings);
        Assert.Contains(result.Report.Diagnostics, d => d.Code == "IWORK_NUMBERS_EXCEL_DESTINATION_UNSUPPORTED");
        Assert.Equal("Preview", Assert.Single(result.Value.Sheets).Name);
    }
}
