using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using M = DocumentFormat.OpenXml.Math;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void TransformTextCasePreservesRunFormatting() {
        using WordDocument document = WordDocument.Create(Path.Combine(_directoryWithFiles, "TextCase.docx"));
        WordParagraph run = document.AddParagraph("Styled");
        run.SetBold().SetItalic().SetUnderline(WordUnderlineStyle.WavyDouble)
            .SetDoubleStrike().SetSuperScript().SetColor(OfficeColor.FromRgb(51, 102, 153)).SetFontFamily("Aptos");

        run.TransformTextCase(OfficeTextCase.ToggleCase);

        WordParagraph actual = document.Paragraphs.Single();
        Assert.Equal("sTYLED", actual.Text);
        Assert.True(actual.Bold);
        Assert.True(actual.Italic);
        Assert.Equal(WordUnderlineStyle.WavyDouble, actual.Underline);
        Assert.True(actual.DoubleStrike);
        Assert.Equal(WordVerticalTextPosition.Superscript, actual.VerticalTextAlignment);
        Assert.Equal("Aptos", actual.FontFamily);
    }

    [Fact]
    public void TransformTextCasePreservesStructuredEquationMarkup() {
        string path = Path.Combine(_directoryWithFiles, "EquationTextCase.docx");
        OfficeMathExpression expression = OfficeMath.Fraction(
            OfficeMath.Text("MIXED Case"),
            OfficeMath.Radical(OfficeMath.Text("OTHER Text")));

        using (WordDocument document = WordDocument.Create(path)) {
            WordParagraph paragraph = document.AddEquation(expression);
            paragraph.TransformTextCase(OfficeTextCase.Lowercase);

            WordEquation equation = Assert.Single(document.Equations);
            Assert.Equal(OfficeMath.Fraction(
                OfficeMath.Text("mixed case"),
                OfficeMath.Radical(OfficeMath.Text("other text"))), equation.ToExpression());
            Assert.Contains("<m:f>", equation.Omml, StringComparison.Ordinal);
            Assert.Contains("<m:rad>", equation.Omml, StringComparison.Ordinal);
            document.Save();
        }

        using WordDocument reopened = WordDocument.Load(path);
        WordEquation actual = Assert.Single(reopened.Equations);
        Assert.Contains("<m:f>", actual.Omml, StringComparison.Ordinal);
        Assert.Contains("<m:rad>", actual.Omml, StringComparison.Ordinal);
        Assert.Equal(OfficeMath.Fraction(
            OfficeMath.Text("mixed case"),
            OfficeMath.Radical(OfficeMath.Text("other text"))), actual.ToExpression());
    }

    [Fact]
    public void TransformTextCasePreservesOmmlRunFormattingInPlace() {
        const string omml = """
            <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
                     xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <m:r><m:rPr><m:nor/></m:rPr><w:rPr><w:b/><w:color w:val="336699"/></w:rPr><m:t>MiXeD</m:t></m:r>
            </m:oMath>
            """;
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph().AddEquation(omml);
        M.Run run = Assert.Single(paragraph.Equation!.MathElement!.Descendants<M.Run>());
        string[] formattingBefore = run.ChildElements
            .Where(element => element.LocalName == "rPr")
            .Select(element => element.OuterXml)
            .ToArray();

        paragraph.TransformTextCase(OfficeTextCase.Lowercase);

        Assert.Equal("mixed", paragraph.Text);
        Assert.Equal(formattingBefore, run.ChildElements
            .Where(element => element.LocalName == "rPr")
            .Select(element => element.OuterXml));
    }

    [Fact]
    public void TransformTextCasePreservesFormattingAcrossStructuredMultiRunContainers() {
        using WordDocument document = WordDocument.Create();

        W.Run Bold(string text) => new W.Run(new W.RunProperties(new W.Bold()), new W.Text(text));
        W.Run Italic(string text) => new W.Run(new W.RunProperties(new W.Italic()), new W.Text(text));
        void AssertRuns(IEnumerable<W.Run> runs, string first, string second) {
            W.Run[] actual = runs.ToArray();
            Assert.Equal(first, actual[0].InnerText);
            Assert.Equal(second, actual[1].InnerText);
            Assert.NotNull(actual[0].RunProperties?.Bold);
            Assert.NotNull(actual[1].RunProperties?.Italic);
        }

        WordParagraph hyperlinkOwner = document.AddParagraph();
        var hyperlink = new W.Hyperlink(Bold("MIXED "), Italic("LINK"));
        hyperlinkOwner._paragraph.Append(hyperlink);
        new WordParagraph(document, hyperlinkOwner._paragraph, hyperlink)
            .TransformTextCase(OfficeTextCase.Lowercase);
        AssertRuns(hyperlink.Elements<W.Run>(), "mixed ", "link");

        WordParagraph fieldOwner = document.AddParagraph();
        var field = new W.SimpleField(Bold("FIELD "), Italic("VALUE")) { Instruction = " AUTHOR " };
        fieldOwner._paragraph.Append(field);
        new WordParagraph(document, fieldOwner._paragraph, field)
            .TransformTextCase(OfficeTextCase.Lowercase);
        AssertRuns(field.Elements<W.Run>(), "field ", "value");

        WordParagraph controlOwner = document.AddParagraph();
        var control = new W.SdtRun(new W.SdtProperties(), new W.SdtContentRun(Bold("CONTROL "), Italic("VALUE")));
        controlOwner._paragraph.Append(control);
        new WordParagraph(document, controlOwner._paragraph, control)
            .TransformTextCase(OfficeTextCase.Lowercase);
        AssertRuns(control.Descendants<W.Run>(), "control ", "value");

        WordParagraph complexOwner = document.AddParagraph();
        var complexRuns = new List<W.Run> {
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode(" AUTHOR ")),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
            Bold("COMPLEX "),
            Italic("VALUE"),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End })
        };
        complexOwner._paragraph.Append(complexRuns);
        new WordParagraph(document, complexOwner._paragraph, complexRuns)
            .TransformTextCase(OfficeTextCase.Lowercase);
        AssertRuns(complexRuns.Skip(3).Take(2), "complex ", "value");
        Assert.Equal(" AUTHOR ", complexRuns[1].InnerText);
    }
}
