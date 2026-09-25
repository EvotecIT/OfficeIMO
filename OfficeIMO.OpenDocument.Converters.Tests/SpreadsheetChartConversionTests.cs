using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class SpreadsheetChartConversionTests {
    [Theory]
    [InlineData("column", ExcelChartType.ColumnClustered)]
    [InlineData("bar", ExcelChartType.BarClustered)]
    [InlineData("line", ExcelChartType.Line)]
    public void ExcelProducedOdsChartConvertsToReopenableXlsx(string kind, ExcelChartType expectedType) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "microsoft-excel-" + kind + "-chart.ods");
        OdsDocument source = OdsDocument.Load(path);
        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;

        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "charts"
            && mapping.Status == OdfConversionMappingStatus.Approximated && mapping.Count == 1);
        Assert.DoesNotContain(result.Report.Mappings, mapping => mapping.Feature == "source-embedded-objects");
        ExcelChart chart = Assert.Single(converted["Data"].Charts);
        Assert.Equal(14.5, converted["Data"].DefaultRowHeight);
        Assert.Equal(8.43, converted["Data"].DefaultColumnWidth);
        Assert.Equal(expectedType, chart.ChartType);
        Assert.Equal("Sales", chart.Title);
        Assert.True(chart.TryGetData(out ExcelChartData data));
        Assert.Equal(new[] { "Jan", "Feb" }, data.Categories);
        Assert.Equal(new[] { 10d, 20d }, Assert.Single(data.Series).Values);

        byte[] xlsx = converted.ToBytes();
        using (SpreadsheetDocument package = SpreadsheetDocument.Open(new MemoryStream(xlsx), false)) {
            Assert.Empty(new OpenXmlValidator().Validate(package));
        }
        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(xlsx));
        Assert.Equal(expectedType, Assert.Single(reopened["Data"].Charts).ChartType);
        Assert.Throws<OdfConversionLossException>(() => source.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss }));
    }

    [Fact]
    public void UnsupportedOdsChartClassRemainsAnExplicitLoss() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            xml.Descendants(chart + "chart").Single().SetAttributeValue(chart + "class", "chart:circle");
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;

        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-embedded-objects"
            && mapping.Status == OdfConversionMappingStatus.Unsupported && mapping.Count == 1);
    }

    [Fact]
    public void BarChartWithoutVerticalAttributeUsesDefaultColumnOrientation() {
        byte[] package = RewritePart("Object 1/styles.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            foreach (XElement properties in xml.Descendants(OdfNamespaces.Style + "chart-properties"))
                properties.Attribute(chart + "vertical")?.Remove();
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        Assert.Equal(ExcelChartType.ColumnClustered, Assert.Single(converted["Data"].Charts).ChartType);
    }

    [Fact]
    public void InvalidBarOrientationRemainsAnExplicitLoss() {
        byte[] package = RewritePart("Object 1/styles.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            xml.Root!.Element(OdfNamespaces.Office + "styles")!.AddFirst(
                new XElement(OdfNamespaces.Style + "default-style",
                    new XAttribute(OdfNamespaces.Style + "family", "chart"),
                    new XElement(OdfNamespaces.Style + "chart-properties",
                        new XAttribute(chart + "vertical", "invalid"))));
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;
        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-embedded-objects"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void ChartDataRespectsTheGlobalExpandedCellBudget() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "microsoft-excel-column-chart.ods");
        OdsDocument source = OdsDocument.Load(path);
        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions { MaximumExpandedCells = 1 });
        using ExcelDocument converted = result.Value;

        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-embedded-objects"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "expansion-limits"
            && mapping.Status == OdfConversionMappingStatus.Skipped);
    }

    [Fact]
    public void ChartDataRespectsTheHiddenWorksheetRowLimit() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "microsoft-excel-column-chart.ods");
        OdsDocument source = OdsDocument.Load(path);
        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions { MaximumRows = 4 });
        using ExcelDocument converted = result.Value;

        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-embedded-objects"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "expansion-limits"
            && mapping.Status == OdfConversionMappingStatus.Skipped);
    }

    [Fact]
    public void ChartSourceRangeBeyondConfiguredRowsReportsExpansionLimit() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XElement categories = xml.Descendants(chart + "categories").Single();
            categories.SetAttributeValue(OdfNamespaces.Table + "cell-range-address", "Data.$A$99:Data.$A$100");
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions { MaximumRows = 50 });
        using ExcelDocument converted = result.Value;
        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "expansion-limits" &&
            mapping.Status == OdfConversionMappingStatus.Skipped);
    }

    [Fact]
    public void SourceNamedRangeCannotReplaceHiddenChartDataOwnerMarker() {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "microsoft-excel-column-chart.ods");
        OdsDocument source = OdsDocument.Load(path);
        source.AddNamedRange("_OfficeIMO_ChartDataOwner", "$'Data'.$A$1");
        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;
        Assert.Single(converted["Data"].Charts);
        Assert.NotNull(converted.GetNamedRange("_OfficeIMO_ChartDataOwner_2"));
        Assert.NotEqual(converted.GetNamedRange("_OfficeIMO_ChartDataOwner_2"),
            converted.GetNamedRange("_OfficeIMO_ChartDataOwner"));
        using ExcelDocument reopened = ExcelDocument.Load(new MemoryStream(converted.ToBytes()));
        Assert.NotNull(reopened.GetNamedRange("_OfficeIMO_ChartDataOwner"));
        Assert.NotNull(reopened.GetNamedRange("_OfficeIMO_ChartDataOwner_2"));
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "named-range-names" &&
            mapping.Status == OdfConversionMappingStatus.Approximated);
    }

    [Fact]
    public void CurrentSheetChartReferencesResolveAgainstHostSheet() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            xml.Descendants(chart + "categories").Single()
                .SetAttributeValue(OdfNamespaces.Table + "cell-range-address", ".$A$2:.$A$3");
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        Assert.Single(converted["Data"].Charts);
    }

    [Fact]
    public void ChartStyleNameMayAlsoExistInAnotherFamily() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XElement chartElement = xml.Descendants(chart + "chart").Single();
            string name = (string)chartElement.Attribute(chart + "style-name")!;
            XElement styles = xml.Root!.Element(OdfNamespaces.Office + "automatic-styles")!;
            styles.AddFirst(new XElement(OdfNamespaces.Style + "style",
                new XAttribute(OdfNamespaces.Style + "name", name),
                new XAttribute(OdfNamespaces.Style + "family", "paragraph")));
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        Assert.Single(converted["Data"].Charts);
    }

    [Fact]
    public void HiddenChartDataWidthRespectsColumnLimit() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XElement series = xml.Descendants(chart + "series").Single();
            series.AddAfterSelf(new XElement(series));
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult(
            new ExcelOpenDocumentConversionOptions { MaximumColumns = 2 });
        using ExcelDocument converted = result.Value;
        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "expansion-limits" &&
            mapping.Status == OdfConversionMappingStatus.Skipped);
    }

    [Theory]
    [InlineData("percentage")]
    [InlineData("currency")]
    public void NumericPercentageAndCurrencySeriesConvert(string valueType) {
        byte[] package = RewritePart("content.xml", xml => {
            XNamespace table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
            XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            XElement[] values = xml.Descendants(table + "table-cell")
                .Where(cell => (string?)cell.Attribute(office + "value-type") == "float").ToArray();
            foreach (XElement cell in values) {
                cell.SetAttributeValue(office + "value-type", valueType);
                if (valueType == "currency") cell.SetAttributeValue(office + "currency", "USD");
            }
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        Assert.Single(converted["Data"].Charts);
    }

    [Fact]
    public void UnlabeledSeriesUsesGeneratedName() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            xml.Descendants(chart + "series").Single().Attribute(chart + "label-cell-address")!.Remove();
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        ExcelChart chart = Assert.Single(converted["Data"].Charts);
        Assert.True(chart.TryGetData(out ExcelChartData data));
        Assert.Equal("Series 1", Assert.Single(data.Series).Name);
    }

    [Fact]
    public void ChartDefaultStyleStackingRemainsExplicitLoss() {
        byte[] package = RewritePart("Object 1/styles.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
            XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            XElement styles = xml.Root!.Element(office + "styles")!;
            styles.AddFirst(new XElement(style + "default-style",
                new XAttribute(style + "family", "chart"),
                new XElement(style + "chart-properties", new XAttribute(chart + "stacked", "true"))));
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        Assert.Empty(converted["Data"].Charts);
    }

    [Fact]
    public void MultilineTitleIsInspectedAndConverted() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
            xml.Descendants(chart + "title").Single().Add(new XElement(text + "p", "Forecast"));
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        Assert.Equal("Sales\nForecast", Assert.Single(source.GetSheet("Data")!.Charts).Title);
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        Assert.Equal("Sales\nForecast", Assert.Single(converted["Data"].Charts).Title);
    }

    [Fact]
    public void ChartTitlePreservesOdfInlineWhitespace() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XNamespace text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
            XElement paragraph = xml.Descendants(chart + "title").Single()
                .Element(text + "p")!;
            paragraph.ReplaceNodes("Net", new XElement(text + "s", new XAttribute(text + "c", "2")), "Sales");
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        Assert.Equal("Net  Sales", Assert.Single(source.GetSheet("Data")!.Charts).Title);
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        Assert.Equal("Net  Sales", Assert.Single(converted["Data"].Charts).Title);
    }

    [Fact]
    public void ReferencedChartTitleRemainsExplicitConversionLoss() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XNamespace table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
            xml.Descendants(chart + "title").Single()
                .SetAttributeValue(table + "cell-range", "Data.$C$1");
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        OdsChart inspected = Assert.Single(source.GetSheet("Data")!.Charts);
        Assert.Equal("Data.$C$1", inspected.TitleCellRangeAddress);
        Assert.Null(inspected.Title);
        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;
        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-embedded-objects"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void ChartClassPrefixAliasesConvertByNamespace() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            xml.Root!.SetAttributeValue(XNamespace.Xmlns + "c", chart.NamespaceName);
            xml.Descendants(chart + "chart").Single().SetAttributeValue(chart + "class", "c:bar");
            foreach (XElement series in xml.Descendants(chart + "series"))
                series.SetAttributeValue(chart + "class", "c:bar");
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        Assert.Equal("chart:bar", Assert.Single(source.GetSheet("Data")!.Charts).ChartClass);
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        Assert.Single(converted["Data"].Charts);
    }

    [Fact]
    public void UnprefixedChartClassUsesDefaultChartNamespace() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XElement chartElement = xml.Descendants(chart + "chart").Single();
            chartElement.SetAttributeValue("xmlns", chart.NamespaceName);
            chartElement.SetAttributeValue(chart + "class", "bar");
            xml.Descendants(chart + "series").Single().SetAttributeValue(chart + "class", "bar");
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        Assert.Equal("chart:bar", Assert.Single(source.GetSheet("Data")!.Charts).ChartClass);
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        Assert.Single(converted["Data"].Charts);
    }

    [Fact]
    public void StackedChartPropertyInSeriesStyleRemainsAnExplicitLoss() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
            XElement seriesStyle = xml.Descendants(style + "style").Single(element =>
                (string?)element.Attribute(style + "name") == "G0S0");
            seriesStyle.Element(style + "chart-properties")!.SetAttributeValue(chart + "stacked", "1");
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;

        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-embedded-objects"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void InheritedStackedChartPropertyRemainsAnExplicitLoss() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
            XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            XElement plot = xml.Descendants(chart + "plot-area").Single();
            string styleName = (string)plot.Attribute(chart + "style-name")!;
            XElement plotStyle = xml.Descendants(style + "style").Single(element =>
                (string?)element.Attribute(style + "name") == styleName);
            plotStyle.Element(style + "chart-properties")!.Attribute(chart + "stacked")?.Remove();
            plotStyle.SetAttributeValue(style + "parent-style-name", "StackedParent");
            xml.Root!.Element(office + "automatic-styles")!.Add(
                new XElement(style + "style",
                    new XAttribute(style + "name", "StackedParent"),
                    new XAttribute(style + "family", "chart"),
                    new XElement(style + "chart-properties", new XAttribute(chart + "stacked", "true"))));
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        Assert.Empty(converted["Data"].Charts);
    }

    [Fact]
    public void ExplicitNonstackedStyleOverridesInheritedStacking() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XNamespace style = "urn:oasis:names:tc:opendocument:xmlns:style:1.0";
            XNamespace office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
            string styleName = (string)xml.Descendants(chart + "plot-area").Single()
                .Attribute(chart + "style-name")!;
            XElement plotStyle = xml.Descendants(style + "style").Single(element =>
                (string?)element.Attribute(style + "name") == styleName);
            plotStyle.Element(style + "chart-properties")!.SetAttributeValue(chart + "stacked", "false");
            plotStyle.SetAttributeValue(style + "parent-style-name", "StackedParent");
            xml.Root!.Element(office + "automatic-styles")!.Add(
                new XElement(style + "style",
                    new XAttribute(style + "name", "StackedParent"),
                    new XAttribute(style + "family", "chart"),
                    new XElement(style + "chart-properties", new XAttribute(chart + "stacked", "true"))));
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        Assert.Single(converted["Data"].Charts);
    }

    [Fact]
    public void SheetLevelShapeChartIsInspectedButNotAssignedAFalseCellAnchor() {
        byte[] package = RewritePart("content.xml", xml => {
            XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
            XNamespace table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
            XElement frame = xml.Descendants(draw + "frame").Single();
            XElement sheet = xml.Descendants(table + "table").Single();
            frame.Remove();
            sheet.Add(new XElement(table + "shapes", frame));
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        OdsChart chart = Assert.Single(source.GetSheet("Data")!.Charts);
        Assert.Null(chart.AnchorRow);
        Assert.Null(chart.AnchorColumn);
        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;
        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-embedded-objects"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    [Fact]
    public void NativeChartInspectionKeepsAllSeriesBeyondTheConversionLimit() {
        byte[] package = RewritePart("Object 1/content.xml", xml => {
            XNamespace chart = "urn:oasis:names:tc:opendocument:xmlns:chart:1.0";
            XElement plot = xml.Descendants(chart + "plot-area").Single();
            XElement series = plot.Elements(chart + "series").Single();
            for (int index = 1; index < 18; index++) series.AddAfterSelf(new XElement(series));
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        Assert.Equal(18, Assert.Single(source.GetSheet("Data")!.Charts).Series.Count);
        using ExcelDocument converted = source.ToExcelDocumentResult().Value;
        Assert.Empty(converted["Data"].Charts);
    }

    [Fact]
    public void ExternalChartObjectHrefIsNotOpenedOrConverted() {
        byte[] package = RewritePart("content.xml", xml => {
            XNamespace draw = "urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
            XNamespace xlink = "http://www.w3.org/1999/xlink";
            xml.Descendants(draw + "object").Single().SetAttributeValue(xlink + "href", "https://example.invalid/chart");
        });
        OdsDocument source = OdsDocument.Load(new MemoryStream(package));
        Assert.Empty(source.GetSheet("Data")!.Charts);
        OdfConversionResult<ExcelDocument> result = source.ToExcelDocumentResult();
        using ExcelDocument converted = result.Value;

        Assert.Empty(converted["Data"].Charts);
        Assert.Contains(result.Report.Mappings, mapping => mapping.Feature == "source-embedded-objects"
            && mapping.Status == OdfConversionMappingStatus.Unsupported);
    }

    private static byte[] RewritePart(string partName, Action<XDocument> edit) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fixtures", "microsoft-excel-column-chart.ods");
        return OdfTestPackageRewriter.Rewrite(File.ReadAllBytes(path), (name, bytes) => {
            if (name != partName) return bytes;
            XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
            edit(xml);
            return Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
        });
    }
}
