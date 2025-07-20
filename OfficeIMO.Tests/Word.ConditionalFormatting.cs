using System.IO;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_TableConditionalFormatting() {
            string filePath = Path.Combine(_directoryWithFiles, "ConditionalFormatting.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(3, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Status";

                table.Rows[1].Cells[0].Paragraphs[0].Text = "Task1";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Done";
                table.Rows[2].Cells[0].Paragraphs[0].Text = "Task2";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Pending";

                table.ConditionalFormatting(
                    "Status",
                    "Done",
                    TextMatchType.Equals,
                    matchFillColorHex: "92d050",
                    noMatchFillColorHex: "ff0000",
                    matchTextFormat: p => p.SetBold(),
                    noMatchTextFormat: p => p.SetUnderline(UnderlineValues.Single));

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTable table = document.Tables[0];
                Assert.Equal("92d050", table.Rows[1].Cells[1].ShadingFillColorHex);
                Assert.Equal("ff0000", table.Rows[2].Cells[1].ShadingFillColorHex);
                Assert.True(table.Rows[1].Cells[1].Paragraphs[0].Bold);
                Assert.Equal(UnderlineValues.Single, table.Rows[2].Cells[1].Paragraphs[0].Underline);
            }
        }

        [Fact]
        public void Test_TableConditionalFormattingAdvanced() {
            string filePath = Path.Combine(_directoryWithFiles, "ConditionalFormattingAdvanced.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(5, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "Name";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "Status";

                table.Rows[1].Cells[0].Paragraphs[0].Text = "Task1";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "Done";

                table.Rows[2].Cells[0].Paragraphs[0].Text = "Task2";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "Pending";

                table.Rows[3].Cells[0].Paragraphs[0].Text = "Task3";
                table.Rows[3].Cells[1].Paragraphs[0].Text = "Skipped";

                table.Rows[4].Cells[0].Paragraphs[0].Text = "Task4";
                table.Rows[4].Cells[1].Paragraphs[0].Text = "Done";

                var builder = table.BeginConditionalFormatting();
                builder.AddRule(
                    "Status",
                    "Done",
                    TextMatchType.Equals,
                    Color.LightGreen,
                    Color.Black,
                    Color.LightPink,
                    Color.Black,
                    highlightColumns: new[] { "Name" },
                    matchTextFormat: p => p.SetBold(),
                    noMatchTextFormat: p => p.SetUnderline(UnderlineValues.Single));

                builder.AddRule(
                    "Status",
                    "Pending",
                    TextMatchType.Equals,
                    Color.Yellow,
                    null,
                    highlightColumns: new[] { "Name" },
                    matchTextFormat: p => p.SetItalic());

                builder.AddRule(
                    new[] {
                        ("Status", "Done", TextMatchType.Equals),
                        ("Name", "Task4", TextMatchType.StartsWith)
                    },
                    matchAll: true,
                    Color.LightSkyBlue,
                    highlightColumns: new[] { "Name" },
                    matchTextFormat: p => {
                        p.SetBold();
                        p.SetUnderline(UnderlineValues.Single);
                    });

                builder.Apply();

                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordTable table = document.Tables[0];
                Assert.Equal("90ee90", table.Rows[1].Cells[1].ShadingFillColorHex);
                Assert.Equal("ffff00", table.Rows[2].Cells[1].ShadingFillColorHex);
                Assert.Equal("ffb6c1", table.Rows[3].Cells[1].ShadingFillColorHex);
                Assert.Equal("90ee90", table.Rows[4].Cells[1].ShadingFillColorHex);

                Assert.Equal("90ee90", table.Rows[1].Cells[0].ShadingFillColorHex);
                Assert.Equal("ffff00", table.Rows[2].Cells[0].ShadingFillColorHex);
                Assert.Equal("ffb6c1", table.Rows[3].Cells[0].ShadingFillColorHex);
                Assert.Equal("87cefa", table.Rows[4].Cells[0].ShadingFillColorHex);

                Assert.True(table.Rows[1].Cells[1].Paragraphs[0].Bold);
                Assert.True(table.Rows[1].Cells[0].Paragraphs[0].Bold);
                Assert.True(table.Rows[2].Cells[1].Paragraphs[0].Italic);
                Assert.Equal(UnderlineValues.Single, table.Rows[3].Cells[1].Paragraphs[0].Underline);
                Assert.True(table.Rows[4].Cells[0].Paragraphs[0].Bold);
                Assert.Equal(UnderlineValues.Single, table.Rows[4].Cells[0].Paragraphs[0].Underline);
            }
        }
    }
}
