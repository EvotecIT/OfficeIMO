using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Charts {
        public static void Example_PieChart(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a pie chart");
            string filePath = System.IO.Path.Combine(folderPath, "PieChart.docx");
            using WordDocument document = WordDocument.Create(filePath);
            var pieChart = document.AddChart("Pie chart");
            pieChart.AddPie("Poland", 15);
            pieChart.AddPie("USA", 30);
            pieChart.AddPie("Brazil", 20);
            document.Save(false);

            var valid = document.ValidateDocument();
            if (valid.Count > 0) {
                Console.WriteLine("Document has validation errors:");
                foreach (var error in valid) {
                    Console.WriteLine(error.Id + ": " + error.Description);
                }
            } else {
                Console.WriteLine("Document is valid.");
            }

            document.Open(openWord);
        }
    }
}
