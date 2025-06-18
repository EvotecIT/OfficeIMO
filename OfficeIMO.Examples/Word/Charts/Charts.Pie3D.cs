using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Charts {
        public static void Example_Pie3DChart(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a 3-D pie chart");
            string filePath = System.IO.Path.Combine(folderPath, "Pie3DChart.docx");
            using WordDocument document = WordDocument.Create(filePath);
            var pie3d = document.AddChart("Pie3D chart");
            pie3d.AddPie3D("Poland", 15);
            pie3d.AddPie3D("USA", 30);
            pie3d.AddPie3D("Brazil", 20);
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
