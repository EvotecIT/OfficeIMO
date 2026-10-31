#region QuickStart
using OfficeIMO.Word;

using var document = WordDocument.Create("report.docx");

document.AddParagraph("Quarterly report").Style = WordParagraphStyles.Heading1;
document.AddParagraph("Created with OfficeIMO.Word.");

var table = document.AddTable(2, 2, WordTableStyle.TableGrid);
table.Rows[0].Cells[0].Paragraphs[0].Text = "Area";
table.Rows[0].Cells[1].Paragraphs[0].Text = "Status";
table.Rows[1].Cells[0].Paragraphs[0].Text = "Documents";
table.Rows[1].Cells[1].Paragraphs[0].Text = "Generated";
table.RepeatHeaderRowAtTheTopOfEachPage = true;
table.Style = WordTableStyle.TableGrid;

document.Save();
#endregion
