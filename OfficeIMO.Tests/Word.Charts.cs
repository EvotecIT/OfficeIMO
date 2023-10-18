using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_BasicWordWithCharts() {
            var filePath = Path.Combine(_directoryWithFiles, "BasicWordWithCharts.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {

                List<string> categories = new List<string>() {
                    "Food", "Housing", "Mix", "Data"
                };

                var paragraphToTest = document.AddParagraph("Test showing adding chart right to existing paragraph");

                // adding charts to document
                document.AddParagraph("This is a bar chart");
                var barChart1 = document.AddBarChart();
                barChart1.AddCategories(categories);
                barChart1.AddChartBar("Brazil", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.Brown);
                barChart1.AddChartBar("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);
                barChart1.AddChartBar("USA", new[] { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                barChart1.BarGrouping = BarGroupingValues.Clustered;
                barChart1.BarDirection = BarDirectionValues.Column;

                Assert.True(barChart1.BarGrouping == BarGroupingValues.Clustered);
                Assert.True(barChart1.BarDirection == BarDirectionValues.Column);
                Assert.True(document.Paragraphs.Count == 3);

                Assert.True(document.Sections[0].Charts.Count == 1);
                Assert.True(document.Charts.Count == 1);

                var lineChart = paragraphToTest.AddLineChart();
                lineChart.AddChartAxisX(categories);
                lineChart.AddChartLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                lineChart.AddChartLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
                lineChart.AddChartLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

                Assert.True(document.Sections[0].Charts.Count == 2);
                Assert.True(document.Charts.Count == 2);

                var paragraph2 = document.AddParagraph("This is a pie chart - but assigned to paragraph");
                var pieChart1 = paragraph2.AddPieChart();
                pieChart1.AddCategories(categories);
                pieChart1.AddChartPie("Poland", new List<int> { 15, 20, 30 });

                Assert.True(document.Sections[0].Charts.Count == 3);
                Assert.True(document.Charts.Count == 3);

                document.AddSection();

                var paragraph4 = document.AddParagraph("Adding a line chart as required 2 - but assigned to paragraph");
                var lineChart4 = paragraph4.AddLineChart();
                lineChart4.AddChartAxisX(categories);
                lineChart4.AddChartLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                lineChart4.AddChartLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
                lineChart4.AddChartLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

                Assert.True(paragraph4.IsChart == false);

                Assert.True(document.Paragraphs[7].IsChart == true);
                Assert.True(document.Paragraphs[7].Chart.RoundedCorners == false);
                lineChart4.RoundedCorners = true;
                Assert.True(document.Paragraphs[7].Chart.RoundedCorners == true);

                document.Paragraphs[7].Chart.RoundedCorners = false;
                Assert.True(document.Paragraphs[7].Chart.RoundedCorners == false);

                Assert.True(lineChart4.RoundedCorners == false);

                Assert.True(document.Sections[0].ParagraphsCharts.Count == 3);
                Assert.True(document.Sections[0].Charts.Count == 3);
                Assert.True(document.Sections[1].Charts.Count == 1);
                Assert.True(document.Sections[1].ParagraphsCharts.Count == 1);
                Assert.True(document.Charts.Count == 4);
                Assert.True(document.ParagraphsCharts.Count == 4);

                var areaChart = document.AddAreaChart("AreaChart");
                areaChart.AddCategories(categories);
               
                areaChart.AddChartArea("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.Brown);
                areaChart.AddChartArea("USA", new List<int>() { 10, 35, 300,13 }, SixLabors.ImageSharp.Color.Green);
                areaChart.AddChartArea("USA", new List<int>() { 10, 35, 230, 150 }, SixLabors.ImageSharp.Color.AliceBlue);
     
                areaChart.AddLegend(LegendPositionValues.Top);


                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {

                Assert.True(document.Sections[0].Charts.Count == 3);
                Assert.True(document.Sections[1].Charts.Count == 2);
                Assert.True(document.Charts.Count == 5);

                document.Save(false);
            }
        }
    }
}
