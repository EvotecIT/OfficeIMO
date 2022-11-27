using System.Collections.Generic;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using SixLabors.ImageSharp;
using VerifyXunit;
using Xunit;

namespace OfficeIMO.VerifyTests.Word;

public class ChartTests : VerifyTestBase {

    private static async Task DoTest(WordprocessingDocument document) {
        var result = ToVerifyResult(document);
        await Verifier.Verify(result, GetSettings());
    }

    [Fact]
    public async Task AddingMultipleCharts() {
        using var document = WordDocument.Create();
        var categories = new List<string> {
            "Food", "Housing", "Mix", "Data"
        };

        document.AddParagraph("This is a bar chart");
        var barChart1 = document.AddBarChart();
        barChart1.AddCategories(categories);
        barChart1.AddChartBar("Brazil", new List<int> { 10, 35, 18, 23 }, Color.Brown);
        barChart1.AddChartBar("Poland", new List<int> { 13, 20, 230, 150 }, Color.Green);
        barChart1.AddChartBar("USA", new[] { 10, 35, 18, 23 }, Color.AliceBlue);
        barChart1.BarGrouping = BarGroupingValues.Clustered;
        barChart1.BarDirection = BarDirectionValues.Column;

        document.AddParagraph("This is a bar chart");
        var barChart2 = document.AddBarChart();
        barChart2.AddCategories(categories);
        barChart2.AddChartBar("USA", 15, Color.Aqua);
        barChart2.RoundedCorners = true;

        document.AddParagraph("This is a pie chart");
        var pieChart = document.AddPieChart();
        pieChart.AddCategories(categories);
        pieChart.AddChartPie("Poland", new List<int> { 15, 20, 30 });

        document.AddParagraph("Adding a line chart as required 1");

        var lineChart = document.AddLineChart();
        lineChart.AddChartAxisX(categories);
        lineChart.AddChartLine("USA", new List<int> { 10, 35, 18, 23 }, Color.AliceBlue);
        lineChart.AddChartLine("Brazil", new List<int> { 10, 35, 300, 18 }, Color.Brown);
        lineChart.AddChartLine("Poland", new List<int> { 13, 20, 230, 150 }, Color.Green);

        document.AddParagraph("Adding a line chart as required 2");

        var lineChart2 = document.AddLineChart();
        lineChart2.AddChartAxisX(categories);
        lineChart2.AddChartLine("USA", new List<int> { 10, 35, 18, 23 }, Color.AliceBlue);
        lineChart2.AddChartLine("Brazil", new List<int> { 10, 35, 300, 18 }, Color.Brown);
        lineChart2.AddChartLine("Poland", new List<int> { 13, 20, 230, 150 }, Color.Green);

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }
    
}
