using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentEndToEnd(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating end-to-end fluent document");
            string filePath = Path.Combine(folderPath, "FluentEndToEnd.docx");
            string imagesPath = Path.Combine(Directory.GetCurrentDirectory(), "Images");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    document.BuiltinDocumentProperties.Title = "Quarterly Review";
                    document.BuiltinDocumentProperties.Creator = "OfficeIMO";
                    document.PageOrientation = OfficePageOrientation.Portrait;
                    document.PageSettings.PageSize = WordPageSize.A4;
                    document.Sections[0].SetMargins(WordMargin.Normal);
                    var paragraph1 = document.AddParagraph();
                    paragraph1 = paragraph1.AddText("Hello ");
                    paragraph1 = paragraph1.AddText("World");
                    paragraph1.SetBold();
                    paragraph1.SetColorHex("#ff0000");
                    paragraph1 = paragraph1.AddText("!");
                    var list2 = document.AddList(WordListStyle.Numbered);
                    int level3 = 0;
                    list2.AddItem("First", level3);
                    list2.AddItem("Second", level3);
                    var table4 = document.AddTable(1, 3);
                    var row5 = table4.Rows[0];
                    row5.Cells[0].Paragraphs[0].Text = System.Convert.ToString("Name") ?? string.Empty;
                    row5.Cells[1].Paragraphs[0].Text = System.Convert.ToString("Role") ?? string.Empty;
                    row5.Cells[2].Paragraphs[0].Text = System.Convert.ToString("Score") ?? string.Empty;
                    row5.RepeatHeaderRowAtTheTopOfEachPage = true;
                    table4.ConditionalFormattingFirstRow = true;
                    var row6 = table4.AddRow();
                    row6.Cells[0].Paragraphs[0].Text = System.Convert.ToString("Alice") ?? string.Empty;
                    row6.Cells[1].Paragraphs[0].Text = System.Convert.ToString("Dev") ?? string.Empty;
                    row6.Cells[2].Paragraphs[0].Text = System.Convert.ToString(98) ?? string.Empty;
                    var row7 = table4.AddRow();
                    row7.Cells[0].Paragraphs[0].Text = System.Convert.ToString("Bob") ?? string.Empty;
                    row7.Cells[1].Paragraphs[0].Text = System.Convert.ToString("Ops") ?? string.Empty;
                    row7.Cells[2].Paragraphs[0].Text = System.Convert.ToString(91) ?? string.Empty;
                    table4.Style = WordTableStyle.TableGrid;
                    table4.Alignment = WordTableAlignment.Center;
                    table4.WidthType = WordTableWidthUnit.Pct;
                    table4.Width = 100 * 50;
                    var image9 = document.AddParagraph().AddImage(Path.Combine(imagesPath, "Kulek.jpg"));
                    image9.Width = OfficeIMO.OfficeLength.Pixels(100);
                    image9.Title = "Chart";
                    image9.Description = "Quarterly chart";
                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}