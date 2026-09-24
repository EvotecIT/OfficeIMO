using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class SpreadsheetConditionalFormattingLossTests {
    [Fact]
    public void DisjointExcelConditionalRangesKeepTheirOwnOrderedMaps() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.AddConditionalRule("A1:A2", ExcelConditionalFormattingOperator.GreaterThan,
            "10", null, "FFFF0000", stopIfTrue: true, priority: 1);
        sheet.AddConditionalRule("A1:A2", ExcelConditionalFormattingOperator.LessThan,
            "0", null, "FF0000FF", priority: 2);
        sheet.AddConditionalRule("C1:C2", ExcelConditionalFormattingOperator.Between,
            "1", "5", "FF00FF00", priority: 3);

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        OdsDocument reopened = OdsDocument.Load(new MemoryStream(result.Value.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        OdsSheet output = reopened.Sheets.Single();
        string? firstStyleName = output.Cell(0, 0).StyleName;
        string? secondStyleName = output.Cell(0, 2).StyleName;
        Assert.NotNull(firstStyleName);
        Assert.NotNull(secondStyleName);
        Assert.NotEqual(firstStyleName, secondStyleName);
        Assert.Equal(firstStyleName, output.Cell(1, 0).StyleName);
        Assert.Equal(secondStyleName, output.Cell(1, 2).StyleName);
        Assert.Equal(new[] { "cell-content()>10", "cell-content()<0" },
            reopened.Styles.Find(OdfStyleFamily.TableCell, firstStyleName!)!.ConditionalMaps.Select(map => map.Condition));
        Assert.Equal("cell-content-is-between(1,5)",
            Assert.Single(reopened.Styles.Find(OdfStyleFamily.TableCell, secondStyleName!)!.ConditionalMaps).Condition);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 3);
        Assert.DoesNotContain(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);

        using ExcelDocument roundTripped = reopened.ToExcelDocument();
        using ExcelDocument reopenedExcel = ExcelDocument.Load(new MemoryStream(roundTripped.ToBytes()));
        ExcelConditionalFormattingInfo[] restored = reopenedExcel.Sheets.Single().GetConditionalFormattingRules().ToArray();
        Assert.Equal(3, restored.Length);
        Assert.Equal(new[] { "FFFF0000", "FF0000FF" }, restored
            .Where(rule => rule.DifferentialFillColorArgb != "FF00FF00")
            .OrderBy(rule => rule.Priority).Select(rule => rule.DifferentialFillColorArgb));
        Assert.All(restored.Where(rule => rule.DifferentialFillColorArgb != "FF00FF00"),
            rule => Assert.Contains("A1", rule.Range));
        Assert.Contains(restored, rule => rule.DifferentialFillColorArgb == "FF00FF00"
            && rule.Range.Contains("C1", StringComparison.Ordinal));
    }

    [Fact]
    public void OverlappingExcelConditionalRangesRemainWholeSheetLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.AddConditionalRule("A1:A2", ExcelConditionalFormattingOperator.GreaterThan,
            "0", fillColor: "FFFF0000");
        sheet.AddConditionalRule("A2:A3", ExcelConditionalFormattingOperator.LessThan,
            "0", fillColor: "FF0000FF");

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Null(result.Value.Sheets.Single().Cell(0, 0).StyleName);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 2);
    }

    [Fact]
    public void UnsupportedDisjointExcelRangeDoesNotPartiallyApplyOtherRange() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.AddConditionalRule("A1:A2", ExcelConditionalFormattingOperator.GreaterThan,
            "0", fillColor: "FFFF0000");
        sheet.CellAt(1, 3).SetValue("Styled").SetBold();
        sheet.AddConditionalRule("C1:C2", ExcelConditionalFormattingOperator.LessThan,
            "0", fillColor: "FF0000FF");

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Null(result.Value.Sheets.Single().Cell(0, 0).StyleName);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 2);
    }

    [Fact]
    public void DisjointExcelConditionalRangesShareOneCellBudget() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.AddConditionalRule("A1:A2048", ExcelConditionalFormattingOperator.GreaterThan,
            "0", fillColor: "FFFF0000");
        sheet.AddConditionalRule("C1:C2049", ExcelConditionalFormattingOperator.LessThan,
            "0", fillColor: "FF0000FF");

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Null(result.Value.Sheets.Single().Cell(0, 0).StyleName);
        Assert.Null(result.Value.Sheets.Single().Cell(0, 2).StyleName);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 2);
    }

    [Fact]
    public void ExcelOrderedNumericFillRulesBecomeOdsStyleMaps() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.AddConditionalRule("A1:A2", ExcelConditionalFormattingOperator.GreaterThan,
            "10", null, "FFFF0000", stopIfTrue: true, priority: 1);
        sheet.AddConditionalRule("A1:A2", ExcelConditionalFormattingOperator.Between,
            "1", "10", "FF00FF00", priority: 2);

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        OdsDocument reopened = OdsDocument.Load(new MemoryStream(result.Value.ToBytes()));
        Assert.True(reopened.Validate().IsValid);
        string? styleName = reopened.Sheets.Single().Cell(0, 0).StyleName;
        Assert.NotNull(styleName);
        Assert.Equal(styleName, reopened.Sheets.Single().Cell(1, 0).StyleName);
        OdfStyle baseStyle = reopened.Styles.Find(OdfStyleFamily.TableCell, styleName!)!;
        Assert.Equal(new[] { "cell-content()>10", "cell-content-is-between(1,10)" },
            baseStyle.ConditionalMaps.Select(map => map.Condition));
        Assert.Equal(new[] { "#FF0000", "#00FF00" }, baseStyle.ConditionalMaps.Select(map =>
            reopened.Styles.Find(OdfStyleFamily.TableCell, map.ApplyStyleName)!.BackgroundColor!.Value.ToString()));
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 2);
        Assert.DoesNotContain(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);

        using ExcelDocument roundTripped = reopened.ToExcelDocument();
        using ExcelDocument reopenedExcel = ExcelDocument.Load(new MemoryStream(roundTripped.ToBytes()));
        ExcelConditionalFormattingInfo[] restored = reopenedExcel.Sheets.Single().GetConditionalFormattingRules()
            .OrderBy(rule => rule.Priority).ToArray();
        Assert.Equal(2, restored.Length);
        Assert.Equal(new[] { "GreaterThan", "Between" }, restored.Select(rule => rule.Operator));
        Assert.Equal(new[] { "FFFF0000", "FF00FF00" }, restored.Select(rule => rule.DifferentialFillColorArgb));
    }

    [Fact]
    public void ExcelRulesWithoutFirstMatchSemanticsRemainExplicitLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.AddConditionalRule("A1:A2", ExcelConditionalFormattingOperator.GreaterThan,
            "10", fillColor: "FFFF0000");
        sheet.AddConditionalRule("A1:A2", ExcelConditionalFormattingOperator.LessThan,
            "0", fillColor: "FF00FF00");

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Null(result.Value.Sheets.Single().Cell(0, 0).StyleName);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 2);
    }

    [Fact]
    public void ExcelConditionalFillDoesNotReplaceAnExistingCellStyle() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.CellAt(1, 1).SetValue("Styled").SetBold();
        sheet.AddConditionalRule("A1", ExcelConditionalFormattingOperator.GreaterThan,
            "0", fillColor: "FFFF0000");

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        OdsSheet output = result.Value.Sheets.Single();
        Assert.NotNull(output.Cell(0, 0).StyleName);
        OdfStyle style = result.Value.Styles.Find(OdfStyleFamily.TableCell, output.Cell(0, 0).StyleName!)!;
        Assert.Empty(style.ConditionalMaps);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Theory]
    [InlineData("1e-29")]
    [InlineData("0.00000000000000000000000000001")]
    public void ExcelNumericBoundsOutsideExactDecimalSubsetRemainExplicitLoss(string bound) {
        using ExcelDocument source = ExcelDocument.Create();
        source.AddWorksheet("Data").AddConditionalRule("A1",
            ExcelConditionalFormattingOperator.GreaterThan, bound, fillColor: "FFFF0000");

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Null(result.Value.Sheets.Single().Cell(0, 0).StyleName);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void ExcelIndexedDifferentialFillRemainsExplicitLoss() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.AddConditionalRule("A1", ExcelConditionalFormattingOperator.GreaterThan,
            "0", fillColor: "FFFF0000");
        using MemoryStream packageStream = new MemoryStream(source.ToBytes());
        using (SpreadsheetDocument package = SpreadsheetDocument.Open(packageStream, true)) {
            Stylesheet stylesheet = package.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            DifferentialFormat format = stylesheet.DifferentialFormats!.Elements<DifferentialFormat>().Single();
            ForegroundColor foreground = format.Fill!.PatternFill!.ForegroundColor!;
            foreground.Rgb = null;
            foreground.Indexed = 10U;
            stylesheet.Save();
        }

        using ExcelDocument imported = ExcelDocument.Load(new MemoryStream(packageStream.ToArray()));
        OdfConversionResult<OdsDocument> result = imported.ToOpenDocumentResult();
        Assert.Null(result.Value.Sheets.Single().Cell(0, 0).StyleName);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void ExcelConditionalRuleIsReportedWhenConvertingToOds() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelSheet sheet = source.AddWorksheet("Data");
        sheet.AddConditionalRule("A1:A3",
            ExcelConditionalFormattingOperator.GreaterThan, "0", fillColor: "FFFF0000");
        sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
            Source = ExcelConditionalFormattingSource.Office2010Extension,
            Range = "A1:A3",
            Type = "Expression",
            Formulas = new[] { "A1<0" },
            DifferentialFillColorArgb = "FFC6EFCE"
        });
        Assert.Equal(2, source.CreateInspectionSnapshot().Worksheets.Single().ConditionalFormattingRuleCount);

        OdfConversionResult<OdsDocument> result = source.ToOpenDocumentResult();
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 2);
    }

    [Theory]
    [InlineData("cell-content()>0", "GreaterThan", "0", null)]
    [InlineData("cell-content()!=0", "NotEqual", "0", null)]
    [InlineData("cell-content-is-between(-1.5, 3)", "Between", "-1.5", "3")]
    [InlineData("cell-content-is-not-between(-1.5, 3)", "NotBetween", "-1.5", "3")]
    public void OdsNumericConditionalFillMapsToExcelRule(string condition, string expectedOperator,
        string expectedFormula1, string? expectedFormula2) {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle highlight = source.Styles.CreateNamed("Highlight", OdfStyleFamily.TableCell);
        highlight.BackgroundColor = OdfColor.Parse("#FFE699");
        OdfStyle ordinary = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        ordinary.AddConditionalMap(condition, highlight.Name, "$'Data'.$A$1");
        sheet.Cell(0, 0).StyleName = ordinary.Name;
        sheet.Cell(1, 0).StyleName = ordinary.Name;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        OdfConversionResult<ExcelDocument> result = reopened.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        using ExcelDocument reopenedExcel = ExcelDocument.Load(new MemoryStream(output.ToBytes()));
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.DoesNotContain(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
        ExcelConditionalFormattingInfo rule = Assert.Single(reopenedExcel.Sheets.Single().GetConditionalFormattingRules());
        Assert.Equal("A1 A2", rule.Range);
        Assert.Equal("CellIs", rule.Type, ignoreCase: true);
        Assert.Equal(expectedOperator, rule.Operator, ignoreCase: true);
        Assert.Equal(expectedFormula2 == null
            ? new[] { expectedFormula1 }
            : new[] { expectedFormula1, expectedFormula2 }, rule.Formulas);
        Assert.Equal("FFFFE699", rule.DifferentialFillColorArgb);
    }

    [Fact]
    public void OrderedOdsConditionalMapsKeepFirstMatchingStyleInExcel() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle first = source.Styles.CreateNamed("First", OdfStyleFamily.TableCell);
        first.BackgroundColor = OdfColor.Parse("#F4CCCC");
        OdfStyle second = source.Styles.CreateNamed("Second", OdfStyleFamily.TableCell);
        second.BackgroundColor = OdfColor.Parse("#D9EAD3");
        OdfStyle ordinary = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        ordinary.AddConditionalMap("cell-content()>0", first.Name);
        ordinary.AddConditionalMap("cell-content()>10", second.Name);
        sheet.Cell(0, 0).StyleName = ordinary.Name;
        sheet.Cell(1, 0).StyleName = ordinary.Name;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Equal(new[] { "cell-content()>0", "cell-content()>10" },
            reopened.Styles.Find(OdfStyleFamily.TableCell, ordinary.Name)!.ConditionalMaps.Select(map => map.Condition));
        OdfConversionResult<ExcelDocument> result = reopened.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        Assert.Equal("0", Assert.Single(output.Sheets.Single().GetConditionalFormattingRules(),
            rule => rule.Priority == 1).Formulas.Single());
        using ExcelDocument reopenedExcel = ExcelDocument.Load(new MemoryStream(output.ToBytes()));
        ExcelConditionalFormattingInfo[] rules = reopenedExcel.Sheets.Single().GetConditionalFormattingRules()
            .OrderBy(rule => rule.Priority).ToArray();

        Assert.Equal(2, rules.Length);
        Assert.Equal(new[] { 1, 2 }, rules.Select(rule => rule.Priority));
        Assert.All(rules, rule => Assert.True(rule.StopIfTrue));
        Assert.All(rules, rule => Assert.Equal("A1 A2", rule.Range));
        Assert.Equal("0", Assert.Single(rules[0].Formulas));
        Assert.Equal("FFF4CCCC", rules[0].DifferentialFillColorArgb);
        Assert.Equal("10", Assert.Single(rules[1].Formulas));
        Assert.Equal("FFD9EAD3", rules[1].DifferentialFillColorArgb);
        Assert.Empty(reopenedExcel.ValidateOpenXml());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 2);
        Assert.DoesNotContain(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void UnsupportedMapInOrderedStyleKeepsTheWholeChainAsLoss() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle fill = source.Styles.CreateNamed("Fill", OdfStyleFamily.TableCell);
        fill.BackgroundColor = OdfColor.Parse("#D9EAD3");
        OdfStyle ordinary = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        ordinary.AddConditionalMap("cell-content-is-whole-number()", fill.Name);
        ordinary.AddConditionalMap("cell-content()>0", fill.Name);
        sheet.Cell(0, 0).StyleName = ordinary.Name;

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        Assert.Empty(output.Sheets.Single().GetConditionalFormattingRules());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 2);
    }

    [Theory]
    [InlineData(16, OdfConversionMappingStatus.Approximated)]
    [InlineData(17, OdfConversionMappingStatus.Unsupported)]
    public void OrderedStyleMapBudgetIsReported(int mapCount, OdfConversionMappingStatus expectedStatus) {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle fill = source.Styles.CreateNamed("Fill", OdfStyleFamily.TableCell);
        fill.BackgroundColor = OdfColor.Parse("#D9EAD3");
        OdfStyle ordinary = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        for (int threshold = 0; threshold < mapCount; threshold++)
            ordinary.AddConditionalMap("cell-content()>" + threshold, fill.Name);
        sheet.Cell(0, 0).StyleName = ordinary.Name;

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        Assert.Equal(mapCount == 16 ? 16 : 0,
            output.Sheets.Single().GetConditionalFormattingRules().Count);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == expectedStatus && mapping.Count == mapCount);
        if (mapCount == 16) {
            using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(output.ToBytes()));
            Assert.Equal(16, reopened.Sheets.Single().GetConditionalFormattingRules().Count);
            Assert.Empty(reopened.ValidateOpenXml());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OdsNumericConditionalFontMapsToExcelDifferentialStyle(bool includeFill) {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle highlight = source.Styles.CreateNamed("Highlight", OdfStyleFamily.TableCell);
        highlight.Color = OdfColor.Parse("#C00000");
        highlight.Bold = true;
        highlight.Italic = true;
        highlight.Underline = true;
        if (includeFill) highlight.BackgroundColor = OdfColor.Parse("#FFE699");
        OdfStyle ordinary = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        ordinary.AddConditionalMap("cell-content()>=2", highlight.Name);
        sheet.Cell(0, 0).StyleName = ordinary.Name;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        OdfConversionResult<ExcelDocument> result = reopened.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        using ExcelDocument reopenedExcel = ExcelDocument.Load(new MemoryStream(output.ToBytes()));

        ExcelConditionalFormattingInfo rule = Assert.Single(reopenedExcel.Sheets.Single().GetConditionalFormattingRules());
        Assert.Equal("A1", rule.Range);
        Assert.Equal("GreaterThanOrEqual", rule.Operator, ignoreCase: true);
        Assert.Equal("2", Assert.Single(rule.Formulas));
        Assert.Equal("FFC00000", rule.DifferentialFontColorArgb);
        Assert.True(rule.DifferentialFontBold);
        Assert.True(rule.DifferentialFontItalic);
        Assert.True(rule.DifferentialFontUnderline);
        Assert.Equal(includeFill ? "FFFFE699" : null, rule.DifferentialFillColorArgb);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Fact]
    public void OdsConditionalFontOffValuesRemainExplicitInExcel() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle baseStyle = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        baseStyle.Bold = true;
        baseStyle.Italic = true;
        baseStyle.Underline = true;
        OdfStyle normal = source.Styles.CreateNamed("Normal", OdfStyleFamily.TableCell);
        normal.Bold = false;
        normal.Italic = false;
        normal.Underline = false;
        baseStyle.AddConditionalMap("cell-content()>0", normal.Name);
        sheet.Cell(0, 0).StyleName = baseStyle.Name;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        OdfConversionResult<ExcelDocument> result = reopened.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        using ExcelDocument reopenedExcel = ExcelDocument.Load(new MemoryStream(output.ToBytes()));

        ExcelConditionalFormattingInfo rule = Assert.Single(reopenedExcel.Sheets.Single().GetConditionalFormattingRules());
        Assert.False(rule.DifferentialFontBold);
        Assert.False(rule.DifferentialFontItalic);
        Assert.False(rule.DifferentialFontUnderline);
        Assert.Null(rule.DifferentialFillColorArgb);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OdsConditionalStrikeRemainsExplicitInExcel(bool strike) {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle baseStyle = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        baseStyle.StrikeThrough = !strike;
        OdfStyle applied = source.Styles.CreateNamed("Strike", OdfStyleFamily.TableCell);
        applied.StrikeThrough = strike;
        baseStyle.AddConditionalMap("cell-content()>0", applied.Name);
        sheet.Cell(0, 0).StyleName = baseStyle.Name;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        OdfConversionResult<ExcelDocument> result = reopened.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        using ExcelDocument reopenedExcel = ExcelDocument.Load(new MemoryStream(output.ToBytes()));

        ExcelConditionalFormattingInfo rule = Assert.Single(reopenedExcel.Sheets.Single().GetConditionalFormattingRules());
        Assert.Equal(strike, rule.DifferentialFontStrike);
        Assert.Null(rule.DifferentialFillColorArgb);
        Assert.Empty(reopenedExcel.ValidateOpenXml());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Theory]
    [InlineData("1pt", 1.0)]
    [InlineData("11.5pt", 11.5)]
    [InlineData("2.54cm", 72.0)]
    [InlineData("409pt", 409.0)]
    public void OdsConditionalFontFamilyAndAbsoluteSizeMapToExcel(string size, double expectedPoints) {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle typography = source.Styles.CreateNamed("Typography", OdfStyleFamily.TableCell);
        typography.FontFamily = "Liberation Serif";
        typography.FontSize = OdfLength.Parse(size);
        typography.Color = OdfColor.Parse("#C00000");
        typography.Bold = true;
        OdfStyle ordinary = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        ordinary.AddConditionalMap("cell-content()>0", typography.Name);
        sheet.Cell(0, 0).StyleName = ordinary.Name;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        OdfConversionResult<ExcelDocument> result = reopened.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        using ExcelDocument reopenedExcel = ExcelDocument.Load(new MemoryStream(output.ToBytes()));

        ExcelConditionalFormattingInfo rule = Assert.Single(reopenedExcel.Sheets.Single().GetConditionalFormattingRules());
        Assert.Equal("Liberation Serif", rule.DifferentialFontName);
        Assert.Equal("FFC00000", rule.DifferentialFontColorArgb);
        Assert.True(rule.DifferentialFontBold);
        Assert.NotNull(rule.DifferentialFontSize);
        Assert.InRange(rule.DifferentialFontSize.Value, expectedPoints - 0.001D, expectedPoints + 0.001D);
        var schemaErrors = reopenedExcel.ValidateOpenXml();
        Assert.True(schemaErrors.Count == 0, string.Join("\n", schemaErrors));
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
    }

    [Theory]
    [InlineData("Liberation Serif, Arial", null)]
    [InlineData(null, "120%")]
    [InlineData(null, "0.5pt")]
    [InlineData(null, "410pt")]
    public void OdsUnrepresentableConditionalTypographyRemainsUnsupported(string? family, string? size) {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle typography = source.Styles.CreateNamed("Typography", OdfStyleFamily.TableCell);
        typography.FontFamily = family;
        typography.FontSize = size == null ? null : OdfLength.Parse(size);
        OdfStyle ordinary = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        ordinary.AddConditionalMap("cell-content()>0", typography.Name);
        sheet.Cell(0, 0).StyleName = ordinary.Name;

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        Assert.Empty(output.Sheets.Single().GetConditionalFormattingRules());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OdsMalformedConditionalFontSizePreservesOtherSupportedStyles(bool includeFill) {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle typography = source.Styles.CreateNamed("Typography", OdfStyleFamily.TableCell);
        typography.FontSize = OdfLength.Points(12);
        if (includeFill) typography.BackgroundColor = OdfColor.Parse("#D9EAD3");
        OdfStyle ordinary = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        ordinary.AddConditionalMap("cell-content()>0", typography.Name);
        sheet.Cell(0, 0).StyleName = ordinary.Name;
        XNamespace styleNamespace = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
        XNamespace foNamespace = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0";
        source.Package.GetXml("styles.xml").Descendants(styleNamespace + "style")
            .Single(element => (string?)element.Attribute(styleNamespace + "name") == typography.Name)
            .Element(styleNamespace + "text-properties")!
            .SetAttributeValue(foNamespace + "font-size", "");
        source.Package.MarkXmlDirty("styles.xml");

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        if (includeFill) {
            ExcelConditionalFormattingInfo rule = Assert.Single(output.Sheets.Single().GetConditionalFormattingRules());
            Assert.Equal("FFD9EAD3", rule.DifferentialFillColorArgb);
            Assert.Null(rule.DifferentialFontSize);
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
                && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        } else {
            Assert.Empty(output.Sheets.Single().GetConditionalFormattingRules());
            Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
                && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        }
    }

    [Fact]
    public void OdsConditionalMapWithUnmappedStyleRemainsUnsupported() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle highlight = source.Styles.CreateNamed("Highlight", OdfStyleFamily.TableCell);
        highlight.TextAlign = "center";
        OdfStyle ordinary = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        ordinary.AddConditionalMap("cell-content()>0", highlight.Name);
        sheet.Cell(0, 0).StyleName = ordinary.Name;

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Empty(output.Sheets.Single().GetConditionalFormattingRules());
    }

    [Fact]
    public void OdsConditionalDoubleStrikeRemainsUnsupported() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle applied = source.Styles.CreateNamed("DoubleStrike", OdfStyleFamily.TableCell);
        applied.StrikeThrough = true;
        applied.LineThroughType = OdfTextDecorationType.Double;
        OdfStyle baseStyle = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        baseStyle.AddConditionalMap("cell-content()>0", applied.Name);
        sheet.Cell(0, 0).StyleName = baseStyle.Name;

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        Assert.Empty(output.Sheets.Single().GetConditionalFormattingRules());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void OdsReversedNumericRangeRemainsAnExplicitLoss() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle fill = source.Styles.CreateNamed("Fill", OdfStyleFamily.TableCell);
        fill.BackgroundColor = OdfColor.Parse("#D9EAD3");
        OdfStyle mapped = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        mapped.AddConditionalMap("cell-content-is-between(3,1)", fill.Name);
        sheet.Cell(0, 0).StyleName = mapped.Name;

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        Assert.Empty(output.Sheets.Single().GetConditionalFormattingRules());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void OdsMixedConditionalMapsReportOnlyTheUnmappedRuleAsUnsupported() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle fill = source.Styles.CreateNamed("Fill", OdfStyleFamily.TableCell);
        fill.BackgroundColor = OdfColor.Parse("#D9EAD3");
        OdfStyle mapped = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        mapped.AddConditionalMap("cell-content()<=-1.5", fill.Name);
        OdfStyle unmodeled = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        unmodeled.AddConditionalMap("cell-content-is-whole-number()", fill.Name);
        sheet.Cell(0, 0).StyleName = mapped.Name;
        sheet.Cell(0, 1).StyleName = unmodeled.Name;

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;

        ExcelConditionalFormattingInfo rule = Assert.Single(output.Sheets.Single().GetConditionalFormattingRules());
        Assert.Equal("A1", rule.Range);
        Assert.Equal("LessThanOrEqual", rule.Operator, ignoreCase: true);
        Assert.Equal("-1.5", Assert.Single(rule.Formulas));
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void RepeatedConditionalCellsAtRuleBudgetRemainMappedAfterSave() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle fill = source.Styles.CreateNamed("Fill", OdfStyleFamily.TableCell);
        fill.BackgroundColor = OdfColor.Parse("#D9EAD3");
        OdfStyle mapped = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        mapped.AddConditionalMap("cell-content()>0", fill.Name);
        sheet.Cell(0, 0).StyleName = mapped.Name;
        XNamespace tableNamespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        source.Package.GetXml("content.xml").Descendants(tableNamespace + "table-cell").Single()
            .SetAttributeValue(tableNamespace + "number-columns-repeated", 4096);
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(output.ToBytes()));

        ExcelConditionalFormattingInfo rule = Assert.Single(reopened.Sheets.Single().GetConditionalFormattingRules());
        Assert.Equal(4096, rule.Range.Split(' ').Length);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.DoesNotContain(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting-cell-limits");
    }

    [Fact]
    public void RepeatedConditionalCellsBeyondRuleBudgetRemainUnsupported() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle fill = source.Styles.CreateNamed("Fill", OdfStyleFamily.TableCell);
        fill.BackgroundColor = OdfColor.Parse("#D9EAD3");
        OdfStyle mapped = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        mapped.AddConditionalMap("cell-content()>0", fill.Name);
        sheet.Cell(0, 0).StyleName = mapped.Name;
        XNamespace tableNamespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        source.Package.GetXml("content.xml").Descendants(tableNamespace + "table-cell").Single()
            .SetAttributeValue(tableNamespace + "number-columns-repeated", 4097);
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;

        Assert.Empty(output.Sheets.Single().GetConditionalFormattingRules());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting-cell-limits"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void SharedStyleReportsASeparateLimitWhenOnlyOneSheetExceedsTheBudget() {
        OdsDocument source = OdsDocument.Create();
        OdfStyle fill = source.Styles.CreateNamed("Fill", OdfStyleFamily.TableCell);
        fill.BackgroundColor = OdfColor.Parse("#D9EAD3");
        OdfStyle mapped = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        mapped.AddConditionalMap("cell-content()>0", fill.Name);
        source.AddSheet("Small").Cell(0, 0).StyleName = mapped.Name;
        source.AddSheet("Large").Cell(0, 0).StyleName = mapped.Name;
        XNamespace tableNamespace = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        source.Package.GetXml("content.xml").Descendants(tableNamespace + "table-cell").Last()
            .SetAttributeValue(tableNamespace + "number-columns-repeated", 4097);
        source.Package.MarkXmlDirty("content.xml");

        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;

        Assert.Single(output.Sheets[0].GetConditionalFormattingRules());
        Assert.Empty(output.Sheets[1].GetConditionalFormattingRules());
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "conditional-formatting-cell-limits"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }
}
