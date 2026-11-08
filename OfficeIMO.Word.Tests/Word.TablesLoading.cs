using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_LoadingWordDocumentCreatedByOffice() {
            string filePath = Path.Combine(_directoryDocuments, "DocumentWithTables.docx");
            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.NotNull(document._document?.MainDocumentPart?.StyleDefinitionsPart);

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
