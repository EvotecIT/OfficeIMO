using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word.Charts {
    internal static class BarCharts {
        public static void Example_AddingBarChart(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with Bar Chart");
            string filePath = System.IO.Path.Combine(folderPath, "Bar Chart Document.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {



                document.AddParagraph("This is a bar chart");

                document.AddParagraph();


                var barChart1 = document.AddBarChart();
                //barChart1.BarGrouping = BarGroupingValues.Clustered;

                //barChart1.BarDirection = BarDirectionValues.Column;

                // var barChart2 = document.AddBarChart();

                //document.AddLineChart();

                //document.AddPieChart();

                //document.AddBarChart3D();

                //barChart2.RoundedCorners = true;
                //Console.WriteLine(barChart2._id);

                document.Save(openWord);
            }
        }
    }
}
