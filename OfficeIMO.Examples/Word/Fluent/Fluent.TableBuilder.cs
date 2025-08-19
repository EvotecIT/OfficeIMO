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
                        .Columns(3).PreferredWidth(Percent: 100)
                        .Header("Name", "Role", "Score")
                        .Row("Alice", "Dev", 98)
                        .Row("Bob", "Ops", 91)
                        .Style(WordTableStyle.TableGrid)
                        .Align(WordHorizontalAlignmentValues.Center))
                    .Table(t => t
                        .From2D(new object[,] {
                            { "Q", "Revenue", "Churn" },
                            { "Q1", "1.1M", "2.1%" },
                            { "Q2", "1.3M", "1.8%" }
                        }).HeaderRow(0))
                    .Table(t => t.AddTable(2, 2).Table!.Rows[0].Cells[0].AddParagraph("TopLeft"))
                    .End()
                    .Save(false);
            }
            Helpers.Open(filePath, openWord);
        }
    }
}

