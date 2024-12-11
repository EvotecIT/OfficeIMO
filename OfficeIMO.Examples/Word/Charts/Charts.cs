using DocumentFormat.OpenXml.Drawing.Charts;

using OfficeIMO.Word;

using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static class Charts {
        public static void Example_AddingMultipleCharts(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with charts");
            string filePath = System.IO.Path.Combine(folderPath, "Charts Document2.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                List<string> categories = new List<string>() { "Food", "Housing", "Mix", "Data" };

                var paragraphToTest = document.AddParagraph("Test showing adding chart right to existing paragraph");

                document.AddParagraph("This is a bar chart 1");
                var barChart1 = document.AddChart("New title");
                barChart1.AddCategories(categories);
                barChart1.AddBar("Brazil", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.Brown);
                barChart1.AddBar("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);
                barChart1.AddBar("USA", new[] { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                barChart1.BarGrouping = BarGroupingValues.Clustered;
                barChart1.BarDirection = BarDirectionValues.Column;

                Console.WriteLine("Title: " + barChart1.Title);

                barChart1.Title = "New title 2";
                Console.WriteLine("Title: " + barChart1.Title);

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                document.AddParagraph("This is a bar chart 2");
                var barChart2 = document.AddChart();
                barChart2.AddCategories(categories);
                barChart2.AddBar("USA", 15, Color.Aqua);
                barChart2.AddBar("Poland", 11, Color.Blue);
                barChart2.RoundedCorners = true;

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                document.AddParagraph("This is a pie chart with 2 pies");
                var pieChart2 = document.AddChart("Test");
                pieChart2.AddPie("Poland", 15);
                pieChart2.AddPie("USA", 30);

                document.AddParagraph("This is a pie chart with 3 pies");
                document.AddChart("Test")
                    .AddPie("Poland", 15)
                    .AddPie("USA", 30)
                    .AddPie("Brazil", 20.2).SetTitle("new title");

                var paragraphForChart = document.AddParagraph("This is a pie chart - but assigned to paragraph");
                var pieChart1 = paragraphForChart.AddChart().SetTitle("My super title");
                pieChart1.AddPie("Poland", 1);
                pieChart1.AddPie("Poland", 10);
                pieChart1.AddPie("Poland", 20);

                document.AddParagraph("Adding a line chart as required 1");

                var lineChart = document.AddChart();
                lineChart.AddChartAxisX(categories);
                lineChart.AddLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                lineChart.AddLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
                lineChart.AddLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                document.AddParagraph("Adding a line chart as required 2");

                var lineChart2 = document.AddChart();
                lineChart2.AddChartAxisX(categories);
                lineChart2.AddLine("USA", new List<int>() { 10, 35, 50, 50 }, SixLabors.ImageSharp.Color.AliceBlue);
                lineChart2.AddLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
                lineChart2.AddLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                // adding charts to paragraphs directly
                var paragraph = document.AddParagraph("This is a bar chart - but assigned to paragraph 1");
                var barChart3 = paragraph.AddChart();
                barChart3.AddCategories(categories);
                barChart3.AddBar("Brazil", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.Brown);
                barChart3.AddBar("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);
                barChart3.AddBar("USA", new[] { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                barChart3.BarGrouping = BarGroupingValues.Clustered;
                barChart3.BarDirection = BarDirectionValues.Column;

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                var paragraph1 = document.AddParagraph("This is a bar chart - but assigned to paragraph 2");
                var barChart5 = paragraph1.AddChart();
                barChart5.AddCategories(categories);
                barChart5.AddBar("USA", 15, Color.Aqua);
                barChart5.RoundedCorners = true;

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                var paragraph2 = document.AddParagraph("This is a pie chart - but assigned to paragraph");
                var pieChart4 = paragraph2.AddChart();
                pieChart4.AddCategories(categories);
                pieChart4.AddPie("Poland", 15);
                pieChart4.AddPie("USA", 18);
                pieChart4.AddPie("Brazil", 10);

                var paragraph3 = document.AddParagraph("Adding a line chart as required 1 - but assigned to paragraph");
                var lineChart3 = paragraph3.AddChart();
                lineChart3.AddChartAxisX(categories);
                lineChart3.AddLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                lineChart3.AddLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
                lineChart3.AddLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                var paragraph4 = document.AddParagraph("Adding a line chart as required 2 - but assigned to paragraph");
                var lineChart4 = paragraph4.AddChart();
                lineChart4.AddChartAxisX(categories);
                lineChart4.AddLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                lineChart4.AddLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
                lineChart4.AddLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                // let's add chart to first paragraph
                var lineChart5 = paragraphToTest.AddChart();
                lineChart5.AddChartAxisX(categories);
                lineChart5.AddLine("USA", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                lineChart5.AddLine("Brazil", new List<int>() { 10, 35, 300, 18 }, SixLabors.ImageSharp.Color.Brown);
                lineChart5.AddLine("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);

                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);

                var table = document.AddTable(3, 3);
                var barChart4 = table.Rows[0].Cells[0].Paragraphs[0].AddChart();
                barChart4.AddCategories(categories);
                barChart4.AddBar("Brazil", new List<int>() { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.Brown);
                barChart4.AddBar("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);
                barChart4.AddBar("USA", new[] { 10, 35, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                barChart4.BarGrouping = BarGroupingValues.Clustered;
                barChart4.BarDirection = BarDirectionValues.Column;

                var areaChart = document.AddChart("AreaChart");
                areaChart.AddCategories(categories);
                areaChart.AddArea("Brazil", new List<int>() { 100, 1, 18, 230 }, SixLabors.ImageSharp.Color.Brown);
                areaChart.AddArea("Poland", new List<int>() { 13, 20, 230, 150 }, SixLabors.ImageSharp.Color.Green);
                areaChart.AddArea("USA", new List<int>() { 10, 305, 18, 23 }, SixLabors.ImageSharp.Color.AliceBlue);
                areaChart.AddLegend(LegendPositionValues.Top);


                Console.WriteLine("Charts count: " + document.Sections[0].Charts.Count);
                Console.WriteLine("Images count: " + document.Sections[0].Images.Count);

                document.Save(openWord);
            }
        }
    }
}
