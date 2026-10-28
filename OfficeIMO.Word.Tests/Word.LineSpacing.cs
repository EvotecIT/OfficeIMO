using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Tests;

/// <summary>
/// Tests adjusting line spacing settings on paragraphs.
/// </summary>
public partial class Word {
    /// <summary>
    /// Creates a document and verifies various line spacing options.
    /// </summary>
    [Fact]
    public void Test_CreatingWordDocumentWithLineRules() {
        var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithLineRules.docx");
        using (var document = WordDocument.Create(filePath)) {
            var par00 = document.AddParagraph("My declaration 1");
            par00.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            par00.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

            var par01 = document.AddParagraph("My declaration 2");
            par01.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(2.5D);
            par01.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero;

            var par02 = document.AddParagraph("My declaration 3");
            par02.LineSpacing = 360;
            par02.LineSpacingRule = WordLineSpacingRule.Exact;


            Assert.Equal(0, checked((int?)(par00.LineSpacingAfter?.ToTwips())));
            Assert.Equal(0, checked((int?)(par00.LineSpacingBefore?.ToTwips())));

            Assert.Equal(50, checked((int?)(par01.LineSpacingAfter?.ToTwips())));
            Assert.Equal(0, checked((int?)(par01.LineSpacingBefore?.ToTwips())));

            Assert.Equal(360, par02.LineSpacing);
            Assert.Equal(WordLineSpacingRule.Exact, par02.LineSpacingRule);

            Assert.Equal(0, checked((int?)(document.Paragraphs[0].LineSpacingAfter?.ToTwips())));
            Assert.Equal(0, checked((int?)(document.Paragraphs[0].LineSpacingBefore?.ToTwips())));

            Assert.Equal(50, checked((int?)(document.Paragraphs[1].LineSpacingAfter?.ToTwips())));
            Assert.Equal(0, checked((int?)(document.Paragraphs[1].LineSpacingBefore?.ToTwips())));

            Assert.Equal(360, document.Paragraphs[2].LineSpacing);
            Assert.Equal(WordLineSpacingRule.Exact, document.Paragraphs[2].LineSpacingRule);
            Assert.True(document.Paragraphs[2].Text == "My declaration 3");

            Assert.Equal(3, document.Paragraphs.Count);

            document.Save();
        }

        using (var document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithLineRules.docx"))) {
            Assert.Equal(0, checked((int?)(document.Paragraphs[0].LineSpacingAfter?.ToTwips())));
            Assert.Equal(0, checked((int?)(document.Paragraphs[0].LineSpacingBefore?.ToTwips())));

            Assert.Equal(50, checked((int?)(document.Paragraphs[1].LineSpacingAfter?.ToTwips())));
            Assert.Equal(0, checked((int?)(document.Paragraphs[1].LineSpacingBefore?.ToTwips())));

            Assert.Equal(360, document.Paragraphs[2].LineSpacing);
            Assert.Equal(WordLineSpacingRule.Exact, document.Paragraphs[2].LineSpacingRule);
            Assert.True(document.Paragraphs[2].Text == "My declaration 3");

            var par02 = document.AddParagraph("My declaration 4");
            par02.LineSpacing = 250;
            par02.LineSpacingRule = WordLineSpacingRule.AtLeast;

            document.Save();
        }

        using (var document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithLineRules.docx"))) {
            Assert.Equal(0, checked((int?)(document.Paragraphs[0].LineSpacingAfter?.ToTwips())));
            Assert.Equal(0, checked((int?)(document.Paragraphs[0].LineSpacingBefore?.ToTwips())));

            Assert.Equal(50, checked((int?)(document.Paragraphs[1].LineSpacingAfter?.ToTwips())));
            Assert.Equal(0, checked((int?)(document.Paragraphs[1].LineSpacingBefore?.ToTwips())));

            Assert.Equal(360, document.Paragraphs[2].LineSpacing);
            Assert.Equal(WordLineSpacingRule.Exact, document.Paragraphs[2].LineSpacingRule);
            Assert.True(document.Paragraphs[2].Text == "My declaration 3");

            Assert.Equal(250, document.Paragraphs[3].LineSpacing);
            Assert.Equal(WordLineSpacingRule.AtLeast, document.Paragraphs[3].LineSpacingRule);
            Assert.True(document.Paragraphs[3].Text == "My declaration 4");

        }
    }
}
