using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class SpreadsheetConditionalFormattingLossTests {
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

    [Fact]
    public void OdsNumericConditionalFillMapsToExcelRule() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle highlight = source.Styles.CreateNamed("Highlight", OdfStyleFamily.TableCell);
        highlight.BackgroundColor = OdfColor.Parse("#FFE699");
        OdfStyle ordinary = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        ordinary.AddConditionalMap("cell-content()>0", highlight.Name, "$'Data'.$A$1");
        sheet.Cell(0, 0).StyleName = ordinary.Name;
        sheet.Cell(1, 0).StyleName = ordinary.Name;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        OdfConversionResult<ExcelDocument> result = reopened.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.DoesNotContain(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
        ExcelConditionalFormattingInfo rule = Assert.Single(output.Sheets.Single().GetConditionalFormattingRules());
        Assert.Equal("A1 A2", rule.Range);
        Assert.Equal("CellIs", rule.Type, ignoreCase: true);
        Assert.Equal("GreaterThan", rule.Operator, ignoreCase: true);
        Assert.Equal("0", Assert.Single(rule.Formulas));
        Assert.Equal("FFFFE699", rule.DifferentialFillColorArgb);
    }

    [Fact]
    public void OdsConditionalMapWithoutFillRemainsUnsupported() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle highlight = source.Styles.CreateNamed("Highlight", OdfStyleFamily.TableCell);
        highlight.Bold = true;
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
    public void OdsMixedConditionalMapsReportOnlyTheUnmappedRuleAsUnsupported() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle fill = source.Styles.CreateNamed("Fill", OdfStyleFamily.TableCell);
        fill.BackgroundColor = OdfColor.Parse("#D9EAD3");
        OdfStyle mapped = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        mapped.AddConditionalMap("cell-content()<=-1.5", fill.Name);
        OdfStyle unmodeled = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        unmodeled.AddConditionalMap("cell-content-is-between(1,3)", fill.Name);
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
