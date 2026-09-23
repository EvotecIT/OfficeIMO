using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using System.IO;
using System.Linq;
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
    public void OdsConditionalStyleMapIsReportedWhenConvertingToExcel() {
        OdsDocument source = OdsDocument.Create();
        OdsSheet sheet = source.AddSheet("Data");
        OdfStyle highlight = source.Styles.CreateNamed("Highlight", OdfStyleFamily.TableCell);
        highlight.BackgroundColor = OdfColor.Parse("#FFE699");
        OdfStyle ordinary = source.Styles.CreateAutomatic(OdfStyleFamily.TableCell);
        ordinary.AddConditionalMap("cell-content()>0", highlight.Name, "$'Data'.$A$1");
        sheet.Cell(0, 0).StyleName = ordinary.Name;

        OdsDocument reopened = OdsDocument.Load(new MemoryStream(source.ToBytes()));
        OdfConversionResult<ExcelDocument> result = reopened.ToExcelDocumentResult();
        using ExcelDocument output = result.Value;
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-conditional-style-maps"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
        Assert.Empty(output.Sheets.Single().GetConditionalFormattingRules());
    }
}
