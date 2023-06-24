using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_CreatingWordDocumentWithLineRules() {
        var filePath = Path.Combine(_directoryWithFiles, "CreatedDocumentWithLineRules.docx");
        using (var document = WordDocument.Create(filePath)) {
            var par00 = document.AddParagraph("My declaration 1");
            par00.LineSpacingAfter = 0;
            par00.LineSpacingBefore = 0;

            var par01 = document.AddParagraph("My declaration 2");
            par01.LineSpacingAfter = 50;
            par01.LineSpacingBefore = 0;

            var par02 = document.AddParagraph("My declaration 3");
            par02.LineSpacing = 360;
            par02.LineSpacingRule = LineSpacingRuleValues.Exact;


            Assert.Equal(0, par00.LineSpacingAfter);
            Assert.Equal(0, par00.LineSpacingBefore);

            Assert.Equal(50, par01.LineSpacingAfter);
            Assert.Equal(0, par01.LineSpacingBefore);

            Assert.Equal(360, par02.LineSpacing);
            Assert.Equal(LineSpacingRuleValues.Exact, par02.LineSpacingRule);

            Assert.Equal(0, document.Paragraphs[0].LineSpacingAfter);
            Assert.Equal(0, document.Paragraphs[0].LineSpacingBefore);

            Assert.Equal(50, document.Paragraphs[1].LineSpacingAfter);
            Assert.Equal(0, document.Paragraphs[1].LineSpacingBefore);

            Assert.Equal(360, document.Paragraphs[2].LineSpacing);
            Assert.Equal(LineSpacingRuleValues.Exact, document.Paragraphs[2].LineSpacingRule);
            Assert.True(document.Paragraphs[2].Text == "My declaration 3");

            Assert.Equal(3, document.Paragraphs.Count);

            document.Save(false);
        }

        using (var document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithLineRules.docx"))) {
            Assert.Equal(0, document.Paragraphs[0].LineSpacingAfter);
            Assert.Equal(0, document.Paragraphs[0].LineSpacingBefore);

            Assert.Equal(50, document.Paragraphs[1].LineSpacingAfter);
            Assert.Equal(0, document.Paragraphs[1].LineSpacingBefore);

            Assert.Equal(360, document.Paragraphs[2].LineSpacing);
            Assert.Equal(LineSpacingRuleValues.Exact, document.Paragraphs[2].LineSpacingRule);
            Assert.True(document.Paragraphs[2].Text == "My declaration 3");

            var par02 = document.AddParagraph("My declaration 4");
            par02.LineSpacing = 250;
            par02.LineSpacingRule = LineSpacingRuleValues.AtLeast;

            document.Save(false);
        }

        using (var document = WordDocument.Load(Path.Combine(_directoryWithFiles, "CreatedDocumentWithLineRules.docx"))) {
            Assert.Equal(0, document.Paragraphs[0].LineSpacingAfter);
            Assert.Equal(0, document.Paragraphs[0].LineSpacingBefore);

            Assert.Equal(50, document.Paragraphs[1].LineSpacingAfter);
            Assert.Equal(0, document.Paragraphs[1].LineSpacingBefore);

            Assert.Equal(360, document.Paragraphs[2].LineSpacing);
            Assert.Equal(LineSpacingRuleValues.Exact, document.Paragraphs[2].LineSpacingRule);
            Assert.True(document.Paragraphs[2].Text == "My declaration 3");

            Assert.Equal(250, document.Paragraphs[3].LineSpacing);
            Assert.Equal(LineSpacingRuleValues.AtLeast, document.Paragraphs[3].LineSpacingRule);
            Assert.True(document.Paragraphs[3].Text == "My declaration 4");

        }
    }
}
