using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static class Charts {
        public static void Example_AddingMultipleCharts(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with charts");
            string filePath = System.IO.Path.Combine(folderPath, "Charts Document.docx");

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

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                document.AddParagraph("This is a bar chart");
                var barChart2 = document.AddBarChart();
                barChart2.AddCategories(categories);
                barChart2.AddChartBar("USA", 15, Color.Aqua);
                barChart2.RoundedCorners = true;


                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                document.AddParagraph("This is a pie chart");
                var pieChart = document.AddPieChart();
                pieChart.AddCategories(categories);
                pieChart.AddChartPie("Poland", new List<int> { 15, 20, 30 });

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);


                document.AddParagraph("Adding a line chart as required 1");

                var lineChart = document.AddLineChart();
                lineChart.AddChartAxisX(categories);
                lineChart.AddChartLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                lineChart.AddChartLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
                lineChart.AddChartLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                document.AddParagraph("Adding a line chart as required 2");

                var lineChart2 = document.AddLineChart();
                lineChart2.AddChartAxisX(categories);
                lineChart2.AddChartLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                lineChart2.AddChartLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
                lineChart2.AddChartLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                // adding charts to paragraphs directly
                var paragraph = document.AddParagraph("This is a bar chart - but assigned to paragraph 1");
                var barChart3 = paragraph.AddBarChart();
                barChart3.AddCategories(categories);
                barChart3.AddChartBar("Brazil", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.Brown);
                barChart3.AddChartBar("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);
                barChart3.AddChartBar("USA", new[] { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                barChart3.BarGrouping = BarGroupingValues.Clustered;
                barChart3.BarDirection = BarDirectionValues.Column;

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                var paragraph1 = document.AddParagraph("This is a bar chart - but assigned to paragraph 2");
                var barChart5 = paragraph1.AddBarChart();
                barChart5.AddCategories(categories);
                barChart5.AddChartBar("USA", 15, Color.Aqua);
                barChart5.RoundedCorners = true;

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                var paragraph2 = document.AddParagraph("This is a pie chart - but assigned to paragraph");
                var pieChart1 = paragraph2.AddPieChart();
                pieChart1.AddCategories(categories);
                pieChart1.AddChartPie("Poland", new List<int> { 15, 20, 30 });

                var paragraph3 = document.AddParagraph("Adding a line chart as required 1 - but assigned to paragraph");
                var lineChart3 = paragraph3.AddLineChart();
                lineChart3.AddChartAxisX(categories);
                lineChart3.AddChartLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                lineChart3.AddChartLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
                lineChart3.AddChartLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                var paragraph4 = document.AddParagraph("Adding a line chart as required 2 - but assigned to paragraph");
                var lineChart4 = paragraph4.AddLineChart();
                lineChart4.AddChartAxisX(categories);
                lineChart4.AddChartLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                lineChart4.AddChartLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
                lineChart4.AddChartLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                // lets add chart to first paragraph
                var lineChart5 = paragraphToTest.AddLineChart();
                lineChart5.AddChartAxisX(categories);
                lineChart5.AddChartLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                lineChart5.AddChartLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
                lineChart5.AddChartLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                var table = document.AddTable(3, 3);
                table.Rows[0].Cells[0].Paragraphs[0].AddBarChart();
                barChart3.AddCategories(categories);
                barChart3.AddChartBar("Brazil", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.Brown);
                barChart3.AddChartBar("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);
                barChart3.AddChartBar("USA", new[] { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                barChart3.BarGrouping = BarGroupingValues.Clustered;
                barChart3.BarDirection = BarDirectionValues.Column;

                var areaChart = document.AddAreaChart("AreaChart");
                areaChart.AddCategories(categories);

                areaChart.AddChartArea("Brazil", new List<int>() { 100, 1, 18, 230 }, SixLabors.ImageSharp.Color.Brown);
                areaChart.AddChartArea("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);
                areaChart.AddChartArea("USA", new List<int>() { 10, 305, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);

                areaChart.AddLegend(LegendPositionValues.Top);


                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                Console.WriteLine("Images count: " + document.Sections[0].Images.Count);

                document.Save(openWord);
            }
        }
    }
}
