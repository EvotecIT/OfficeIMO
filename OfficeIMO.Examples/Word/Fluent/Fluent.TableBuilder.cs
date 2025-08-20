using System;
using System.IO;
using OfficeIMO;
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
                        .Columns(3).PreferredWidth(Percent: 100)
                        .Header("Name", "Role", "Score")
                        .Row("Alice", "Dev", 98)
                        .Row("Bob", "Ops", 91)
                        .Style(WordTableStyle.TableGrid)
                        .Align(HorizontalAlignment.Center))
                    .Table(t => t
                        .From2D(new object[,] {
                            { "Q", "Revenue", "Churn" },
                            { "Q1", "1.1M", "2.1%" },
                            { "Q2", "1.3M", "1.8%" }
                        }).HeaderRow(0))
                    .Table(t => t
                        .AddTable(2, 3)
                        .ForEachCell((r, c, cell) => cell.AddParagraph($"R{r}C{c}", true))
                        .Cell(1, 3, cell => cell.AddParagraph("Last", true))
                        .InsertRow(3, "A", "B", "C")
                        .InsertColumn(4, "X", "Y", "Z")
                        .RowStyle(1, r => r.Cells.ForEach(c => c.ShadingFillColorHex = "ffcccc"))
                        .ColumnStyle(2, c => c.ShadingFillColorHex = "ccffcc"))
                    .End()
                    .Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}

