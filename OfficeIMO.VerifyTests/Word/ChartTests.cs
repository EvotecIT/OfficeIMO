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
        var result = await ToVerifyResult(document);
        await Verifier.Verify(result, GetSettings());
    }

    [Fact]
    public async Task AddingMultipleCharts() {
        using var document = WordDocument.Create();
        var categories = new List<string> {
            "Food", "Housing", "Mix", "Data"
        };

        document.AddParagraph("This is a bar chart");
        var barChart1 = document.AddChart();
        barChart1.AddCategories(categories);
        barChart1.AddBar("Brazil", new List<int> { 10, 35, 18, 23 }, Color.Brown);
        barChart1.AddBar("Poland", new List<int> { 13, 20, 230, 150 }, Color.Green);
        barChart1.AddBar("USA", new[] { 10, 35, 18, 23 }, Color.AliceBlue);
        barChart1.BarGrouping = BarGroupingValues.Clustered;
        barChart1.BarDirection = BarDirectionValues.Column;

        document.AddParagraph("This is a bar chart");
        var barChart2 = document.AddChart();
        barChart2.AddCategories(categories);
        barChart2.AddBar("USA", 15, Color.Aqua);
        barChart2.RoundedCorners = true;

        document.AddParagraph("This is a pie chart");
        var pieChart = document.AddChart();
        //pieChart.AddCategories(categories);
        pieChart.AddPie("Poland", 15);
        pieChart.AddPie("USA", 25);
        pieChart.AddPie("Brazil", 60);

        document.AddParagraph("Adding a line chart as required 1");

        var lineChart = document.AddChart();
        lineChart.AddChartAxisX(categories);
        lineChart.AddLine("USA", new List<int> { 10, 35, 18, 23 }, Color.AliceBlue);
        lineChart.AddLine("Brazil", new List<int> { 10, 35, 300, 18 }, Color.Brown);
        lineChart.AddLine("Poland", new List<int> { 13, 20, 230, 150 }, Color.Green);

        document.AddParagraph("Adding a line chart as required 2");

        var lineChart2 = document.AddChart();
        lineChart2.AddChartAxisX(categories);
        lineChart2.AddLine("USA", new List<int> { 10, 35, 18, 23 }, Color.AliceBlue);
        lineChart2.AddLine("Brazil", new List<int> { 10, 35, 300, 18 }, Color.Brown);
        lineChart2.AddLine("Poland", new List<int> { 13, 20, 230, 150 }, Color.Green);

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }

}
