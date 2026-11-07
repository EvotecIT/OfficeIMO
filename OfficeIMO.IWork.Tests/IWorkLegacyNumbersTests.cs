using System.Globalization;
using System.Xml.Linq;
using OfficeIMO.Excel;

namespace OfficeIMO.IWork.Tests;

public sealed class IWorkLegacyNumbersTests {
    [Fact]
    public void Numbers09CellsStylesDimensionsAndMergesSurviveSavedXlsxWithFormulaCachesReported() {
        using MemoryStream package = CreatePackage();
        using IWorkSourceDocument source = IWorkSourceDocument.Load(package);
        IWorkNumbersProjection projection = source.ReadNumbers();
        IWorkTable table = Assert.Single(Assert.Single(projection.Sheets).Tables);
        Assert.Equal(new double?[] { 24, 18, 30 }, Enumerable.Range(1, 3).Select(table.GetRowHeight));
        Assert.Equal(new double?[] { 60, 90, 120, 72 }, Enumerable.Range(1, 4).Select(table.GetColumnWidth));
        Assert.Equal((12d, 20d, 342d, 72d), (table.Geometry!.LeftPoints, table.Geometry.TopPoints,
            table.Geometry.WidthPoints, table.Geometry.HeightPoints));
        Assert.Equal("Selected", table.GetCell(3, 2)!.Value);
        Assert.Equal(IWorkCellKind.Boolean, table.GetCell(1, 4)!.Kind);
        Assert.Equal(true, table.GetCell(1, 4)!.Value);
        Assert.Equal(new DateTime(2001, 1, 2), table.GetCell(2, 3)!.Value);
        Assert.Equal(IWorkCellKind.Duration, table.GetCell(3, 1)!.Kind);
        Assert.Equal(3660d, table.GetCell(3, 1)!.Value);
        IWorkTableCell formula = table.GetCell(2, 4)!;
        Assert.Equal(IWorkCellKind.Formula, formula.Kind);
        Assert.Equal(IWorkCellKind.Number, formula.ValueKind);
        Assert.Equal(42d, formula.Value);
        Assert.True(formula.CachedValueIsComplete);
        Assert.False(formula.FormulaIsComplete);
        Assert.Null(formula.Formula);
        IWorkTableCell blank = table.GetCell(1, 2)!;
        Assert.Equal(IWorkCellKind.Empty, blank.Kind);
        Assert.Equal("336699", blank.Fill!.Color!.RgbHex);
        Assert.Equal(2, blank.Padding!.LeftPoints);
        Assert.Equal(IWorkCellVerticalAlignment.Middle, blank.VerticalAlignment);
        Assert.Equal(IWorkNumberFormatKind.Currency, table.GetCell(2, 1)!.NumberFormat!.Kind);
        Assert.Equal(IWorkNumberFormatKind.Percentage, table.GetCell(2, 2)!.NumberFormat!.Kind);

        using var converted = source.ToExcelResult(new IWorkConversionOptions {
            Mode = IWorkConversionMode.EditableOnly,
            AllowPartialEditableReconstruction = true
        });
        Assert.False(converted.IsVisualFallback);
        Assert.Contains(converted.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_NUMBERS_GRID_DISPLAY_APPROXIMATED"
            && diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Contains(converted.Report.Diagnostics, diagnostic => diagnostic.Code == "IWORK_LEGACY_NUMBERS_FORMULA_CACHE_ONLY"
            && diagnostic.LossKind == OfficeConversionLossKind.Omission);
        using var saved = new MemoryStream();
        converted.Value.Save(saved);
        saved.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(saved);
        ExcelSheet sheet = Assert.Single(reopened.Sheets);
        Assert.Equal("FF336699", sheet.Cell(1, 2).GetStyle().FillColorArgb);
        Assert.True(sheet.Cell(1, 1).GetStyle().Bold);
        Assert.Equal(OfficeLength.Points(14), sheet.Cell(1, 1).GetStyle().FontSize);
        Assert.Equal("\"USD \"#,##0.00;\"USD \"-#,##0.00", sheet.Cell(2, 1).GetStyle().NumberFormatCode);
        Assert.Equal("0.0%", sheet.Cell(2, 2).GetStyle().NumberFormatCode);
        Assert.Equal(1234.5d, sheet.Cell(2, 1).GetValue<double>());
        Assert.Equal(0.125d, sheet.Cell(2, 2).GetValue<double>());
        Assert.Equal(42d, sheet.Cell(2, 4).GetValue<double>());

        saved.Position = 0;
        using var zip = new ZipArchive(saved, ZipArchiveMode.Read, leaveOpen: true);
        using Stream xml = zip.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XDocument worksheet = XDocument.Load(xml);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        Assert.Equal("B1:C1", Assert.Single(worksheet.Descendants(ns + "mergeCell")).Attribute("ref")!.Value);
        Assert.Equal(new[] { 24d, 18d, 30d }, worksheet.Descendants(ns + "row")
            .Select(row => double.Parse(row.Attribute("ht")!.Value, CultureInfo.InvariantCulture)));
        XElement cached = Assert.Single(worksheet.Descendants(ns + "c"), cell => (string?)cell.Attribute("r") == "D2");
        Assert.Null(cached.Element(ns + "f"));
        Assert.Equal("42", cached.Element(ns + "v")!.Value);
        Assert.Equal("1", Assert.Single(worksheet.Descendants(ns + "c"), cell => (string?)cell.Attribute("r") == "D1")
            .Element(ns + "v")!.Value);
        Assert.Equal(4, worksheet.Descendants(ns + "col").Count());
        Assert.Contains(worksheet.Descendants(ns + "col"), column => (string?)column.Attribute("min") == "2"
            && Math.Abs(double.Parse(column.Attribute("width")!.Value, CultureInfo.InvariantCulture)
                - (90d * 96 / 72 - 5) / 7) < 0.000001);
    }

    [Theory]
    [InlineData("cells")]
    [InlineData("dimensions")]
    [InlineData("merges")]
    public void Numbers09RichProjectionRetainsExistingCellDimensionAndMergeLimits(string limit) {
        using MemoryStream package = CreatePackage(additionalMerge: limit == "merges");
        var options = new IWorkLoadOptions();
        if (limit == "cells") options.MaxMaterializedCells = 2;
        if (limit == "dimensions") options.MaxTableDimensionEntries = 6;
        if (limit == "merges") options.MaxTableMergedRanges = 1;
        using IWorkSourceDocument source = IWorkSourceDocument.Load(package, options);
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => source.ReadNumbers());
        Assert.Contains(limit == "cells" ? "cell count" : limit == "dimensions" ? "dimension entries" : "merged-range", error.Message);
    }

    private static MemoryStream CreatePackage(bool additionalMerge = false) {
        string index = """
            <ls:document xmlns:ls="http://developer.apple.com/namespaces/ls"
              xmlns:sf="http://developer.apple.com/namespaces/sf" xmlns:sfa="http://developer.apple.com/namespaces/sfa"
              ls:version="92008102400">
              <ls:stylesheet>
                <sf:cell-style sfa:ID="base-cell" sf:ident="body-cell"><sf:property-map>
                  <sf:SFTCellStylePropertyLayoutStyle><sf:layoutstyle sfa:ID="cell-layout"><sf:property-map>
                    <sf:padding><sf:padding sf:left="2" sf:right="3" sf:top="1" sf:bottom="1"/></sf:padding>
                    <sf:verticalAlignment><sf:number sfa:number="1"/></sf:verticalAlignment>
                  </sf:property-map></sf:layoutstyle></sf:SFTCellStylePropertyLayoutStyle>
                  <sf:SFTCellStylePropertyParagraphStyle><sf:paragraphstyle sfa:ID="cell-paragraph"><sf:property-map>
                    <sf:bold><sf:number sfa:number="1"/></sf:bold>
                    <sf:fontSize><sf:number sfa:number="14"/></sf:fontSize>
                    <sf:fontName><sf:string sfa:string="Arial"/></sf:fontName>
                  </sf:property-map></sf:paragraphstyle></sf:SFTCellStylePropertyParagraphStyle>
                  <sf:fill><sf:color sfa:r="0.2" sfa:g="0.4" sfa:b="0.6" sfa:a="1"/></sf:fill>
                </sf:property-map></sf:cell-style>
                <sf:cell-style sfa:ID="currency-cell" sf:parent-ident="body-cell"><sf:property-map>
                  <sf:SFTCellStylePropertyFormatType><sf:number sfa:number="257"/></sf:SFTCellStylePropertyFormatType>
                  <sf:SFTCellStylePropertyNumberFormat><sf:number-format sf:format-type="1"
                    sf:format-decimal-places="2" sf:format-currency-code="USD" sf:format-negative-style="0"
                    sf:format-show-thousands-separator="true" sf:format-use-accounting-style="false"/>
                  </sf:SFTCellStylePropertyNumberFormat>
                </sf:property-map></sf:cell-style>
                <sf:cell-style sfa:ID="percentage-cell" sf:parent-ident="body-cell"><sf:property-map>
                  <sf:SFTCellStylePropertyFormatType><sf:number sfa:number="258"/></sf:SFTCellStylePropertyFormatType>
                  <sf:SFTCellStylePropertyNumberFormat><sf:number-format sf:format-type="2"
                    sf:format-decimal-places="1" sf:format-negative-style="0" sf:format-show-thousands-separator="false"/>
                  </sf:SFTCellStylePropertyNumberFormat>
                </sf:property-map></sf:cell-style>
              </ls:stylesheet>
              <ls:workspace-array><ls:workspace ls:workspace-name="Migration"><ls:page-info><sf:layers><sf:layer>
                <sf:drawables><sf:tabular-info>
                  <sf:geometry><sf:position sfa:x="12" sfa:y="20"/><sf:size sfa:w="342" sfa:h="72"/></sf:geometry>
                  <sf:tabular-model sf:name="Typed values"><sf:grid>
                    <sf:columns sf:count="4"><sf:grid-column sf:width="60"/><sf:grid-column sf:width="90"/>
                      <sf:grid-column sf:width="120"/><sf:grid-column sf:width="72"/></sf:columns>
                    <sf:rows sf:count="3"><sf:grid-row sf:height="24"/><sf:grid-row sf:height="18"/>
                      <sf:grid-row sf:height="30"/></sf:rows>
                    <sf:datasource>
                      <sf:t sf:s="base-cell" sf:ct="0"><sf:ct sfa:s="Label"/></sf:t>
                      <sf:g sf:s="base-cell" sf:col-span="2"/><sf:s sf:ho="1"/><sf:cb sf:v="true"/>
                      <sf:n sf:s="currency-cell" sf:ct="253" sf:v="1234.5"/>
                      <sf:n sf:s="percentage-cell" sf:v="0.125"/><sf:d sf:cell-date="86400"/>
                      <sf:f><sf:fo sf:fs="=40+2"/><sf:r><sf:rn sf:v="42"/></sf:r></sf:f>
                      <sf:du sf:ct="253" sf:du="3660"/>
                      <sf:pm><sf:menu-choices><sf:t sfa:ID="choice"><sf:ct sfa:s="Selected"/></sf:t></sf:menu-choices>
                        <sf:proxied-cell-ref sfa:IDREF="choice"/></sf:pm>
                      <sf:t sf:s="base-cell"><sf:ct><sf:so><sf:text-body><sf:p>Rich <sf:span>text</sf:span></sf:p>
                        </sf:text-body></sf:so></sf:ct></sf:t>
                      <sf:sl sf:v="7"/>
                    </sf:datasource>
                  </sf:grid></sf:tabular-model>
                </sf:tabular-info></sf:drawables>
              </sf:layer></sf:layers></ls:page-info></ls:workspace></ls:workspace-array>
            </ls:document>
            """;
        if (additionalMerge) index = index.Replace("<sf:t sf:s=\"base-cell\"><sf:ct><sf:so>",
            "<sf:t sf:s=\"base-cell\" sf:col-span=\"2\"><sf:ct><sf:so>")
            .Replace("<sf:sl sf:v=\"7\"/>", "<sf:s sf:ho=\"1\"/>");
        var package = new MemoryStream();
        using (var zip = new ZipArchive(package, ZipArchiveMode.Create, leaveOpen: true))
        using (var writer = new StreamWriter(zip.CreateEntry("index.xml").Open())) writer.Write(index);
        package.Position = 0;
        return package;
    }
}
