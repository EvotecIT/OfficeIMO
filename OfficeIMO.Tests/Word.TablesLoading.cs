using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_LoadingWordDocumentCreatedByOffice() {
            string filePath = Path.Combine(_directoryDocuments, "DocumentWithTables.docx");
            using (WordDocument document = WordDocument.Load(filePath)) {
                // We check for style definition part.
                StyleDefinitionsPart? styleDefinitionsPart = document._document?.MainDocumentPart?.GetPartsOfType<StyleDefinitionsPart>().FirstOrDefault();
                // It should exists
                Assert.NotNull(styleDefinitionsPart);

                // we now check if all table styles are available for use
                List<Style> listTableStyles = new List<Style>();
                var styles = styleDefinitionsPart!.Styles?.OfType<Style>()?.ToList() ?? new List<Style>();
                foreach (var style in styles) {
                    if (style.Type != null && style.Type == StyleValues.Table) {
                        listTableStyles.Add(style);
                    }
                }
                // Style counts may vary between Office versions and template changes.
                // Ensure we have a healthy set of table styles and overall styles.
                Assert.True(listTableStyles.Count >= 80);
                Assert.True(styles.Count >= 100);

                // OfficeIMO settings
                Assert.True(document.Tables.Count == 2);

                var table = document.Tables[0];
                Assert.True(table.Style == WordTableStyle.TableGrid);

                table.Style = WordTableStyle.GridTable1LightAccent4;

                Assert.True(table.Style == WordTableStyle.GridTable1LightAccent4);

                WordTable wordTable = document.AddTable(3, 4, WordTableStyle.GridTable1LightAccent5);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                Assert.True(wordTable.Style == WordTableStyle.GridTable1LightAccent5);

                wordTable = document.AddTable(3, 4, WordTableStyle.GridTable1LightAccent6);
                wordTable.Rows[0].Cells[0].Paragraphs[0].Text = "Test 1";

                Assert.True(wordTable.Style == WordTableStyle.GridTable1LightAccent6);

                Assert.True(document.Tables.Count == 4);
            }
        }
    }
}
