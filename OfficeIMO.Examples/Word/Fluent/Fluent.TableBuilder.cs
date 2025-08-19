using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Fluent;

namespace OfficeIMO.Examples.Word {
    internal static partial class FluentDocument {
        public static void Example_FluentTableBuilder(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with fluent tables");
            string filePath = Path.Combine(folderPath, "FluentTableBuilder.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AsFluent()
                    .Table(t => t
                        .Columns(3).PreferredWidth(percent: 100)
                        .Header("Name", "Role", "Score")
                        .Row("Alice", "Dev", 98)
                        .Row("Bob", "Ops", 91)
                        .Style(WordTableStyle.TableGrid)
                        .Align(HorizontalAlignment.Center))
                    .Table(t => t
                        .From2D(new object[,] {
                            { "Q",  "Revenue", "Churn" },
                            { "Q1", "1.1M",    "2.1%" },
                            { "Q2", "1.3M",    "1.8%" }
                        })
                        .HeaderRow(1))
                    .Table(t => t
                        .Create(rows: 2, cols: 3)
                        .ForEachCell((r, c, cell) => cell.Text($"R{r}C{c}"))
                        .Cell(1, 3).Text("Last")
                        .InsertRow(2, "A", "B", "C")
                        .InsertColumn(2, "X", "Y", "Z")
                        .Row(1).EachCell(c => c.Shading("#ffcccc"))
                        .Column(3).Shading("#ccffcc")
                        .Merge(fromRow: 1, fromCol: 1, toRow: 2, toCol: 2)
                        .DeleteRow(2)
                        .DeleteColumn(2))
                    .End();
                document.Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}
