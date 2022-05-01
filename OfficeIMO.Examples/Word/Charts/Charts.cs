using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word.Charts {
    internal static class Charts {
        public static void Example_AddingMultipleCharts(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with charts");
            string filePath = System.IO.Path.Combine(folderPath, "Charts Document.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("This is a bar chart");
                var barChart1 = document.AddBarChart();
                //barChart1.BarGrouping = BarGroupingValues.Clustered;
                //barChart1.BarDirection = BarDirectionValues.Column;

                document.AddParagraph("This is a pie chart");
                var pieChart = document.AddPieChart();

                //document.AddBarChart3D();

                //barChart2.RoundedCorners = true;
                //Console.WriteLine(barChart2._id);

                document.AddParagraph("This is a line chart");
                var lineChart = document.AddLineChart();


                document.Save(openWord);
            }
        }
    }
}
