using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Rtf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;
using M = DocumentFormat.OpenXml.Math;
using Xunit;

namespace OfficeIMO.Rtf.Tests;

public sealed class RtfEquationTests {
    [Fact]
    public void WordRtfRoundTrip_UsesNativeEqFieldAndPreservesCachedEquationText() {
        const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara>";
        using WordDocument word = WordDocument.Create();
        word.AddEquation(omml);

        RtfConversionResult<RtfDocument> conversion = word.ToRtfDocumentResult();
        RtfField field = Assert.Single(Assert.Single(conversion.Value.Paragraphs).Inlines.OfType<RtfField>());

        Assert.True(field.IsEquation);
        Assert.Equal("(a)/(b)", field.ToPlainText());
        Assert.Contains("\\f(a,b)", field.Instruction, StringComparison.Ordinal);
        Assert.Contains(conversion.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "WordRtfEquationsMappedToEqFields" &&
            diagnostic.Action == RtfConversionAction.Substituted);
        Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic =>
            diagnostic.Code == "WordRtfElementOmitted" && diagnostic.Feature == nameof(WordEquation));

        string rtf = conversion.Value.ToRtf();
        Assert.Contains("\\field", rtf, StringComparison.Ordinal);
        Assert.Contains("EQ", rtf, StringComparison.Ordinal);

        RtfReadResult read = RtfDocument.Read(rtf);
        RtfField readField = Assert.Single(Assert.Single(read.Document.Paragraphs).Inlines.OfType<RtfField>());
        Assert.True(readField.IsEquation);
        Assert.Equal("(a)/(b)", readField.ToPlainText());

        using WordDocument roundTrip = read.Document.ToWordDocument();
        WordEquation equation = Assert.Single(roundTrip.Equations);
        Assert.Equal(WordEquationRepresentation.EquationField, equation.Representation);
        Assert.Equal("(a)/(b)", equation.Text);
        Assert.Contains("\\f(a,b)", equation.FieldInstruction!, StringComparison.Ordinal);
    }

    [Fact]
    public void WordToRtf_MapsOmmlFractionTypeVariantsToNativeEqFields() {
        const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath>" +
            "<m:f><m:fPr><m:type m:val=\"lin\"/></m:fPr><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>" +
            "<m:f><m:fPr><m:type m:val=\"noBar\"/></m:fPr><m:num><m:r><m:t>c</m:t></m:r></m:num><m:den><m:r><m:t>d</m:t></m:r></m:den></m:f>" +
            "<m:f><m:fPr><m:type m:val=\"skw\"/></m:fPr><m:num><m:r><m:t>e</m:t></m:r></m:num><m:den><m:r><m:t>f</m:t></m:r></m:den></m:f>" +
            "</m:oMath></m:oMathPara>";
        using WordDocument word = WordDocument.Create();
        word.AddEquation(omml);

        RtfField field = Assert.Single(Assert.Single(word.ToRtfDocument().Paragraphs).Inlines.OfType<RtfField>());

        Assert.Equal("a/bstack(c,d)e⁄f", field.ToPlainText());
        Assert.Contains("a/b", field.Instruction, StringComparison.Ordinal);
        Assert.Contains("\\a\\co1(c,d)", field.Instruction, StringComparison.Ordinal);
        Assert.Contains("e⁄f", field.Instruction, StringComparison.Ordinal);
    }

    [Fact]
    public void RtfParagraph_AddEquationField_NormalizesEqPrefix() {
        RtfParagraph paragraph = RtfDocument.Create().AddParagraph();

        RtfField field = paragraph.AddEquationField("\\r(,x)", "sqrt(x)");

        Assert.True(field.IsEquation);
        Assert.StartsWith("EQ ", field.Instruction, StringComparison.Ordinal);
        Assert.Equal("sqrt(x)", field.ToPlainText());
    }

    [Fact]
    public void WordToRtf_MapsEquationsInsideVisibleRevisionWrappers() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("Formula: ");
        paragraph._paragraph.Append(
            new InsertedRun(new M.OfficeMath(new M.Run(new M.Text("inserted")))) {
                Id = "1",
                Author = "Reviewer"
            },
            new MoveToRun(new M.OfficeMath(new M.Run(new M.Text("moved")))) {
                Id = "2",
                Author = "Reviewer"
            },
            new InsertedRun(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" EQ \\f(a,b) ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("(a)/(b)")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End })) {
                Id = "3",
                Author = "Reviewer"
            },
            new MoveToRun(
                new SimpleField(new Run(new Text("simple field"))) {
                    Instruction = " EQ \\r(,x) "
                }) {
                Id = "4",
                Author = "Reviewer"
            });

        RtfDocument rtf = word.ToRtfDocument();
        RtfField[] fields = Assert.Single(rtf.Paragraphs).Inlines.OfType<RtfField>().ToArray();

        Assert.Equal(new[] { "inserted", "moved", "(a)/(b)", "simple field" }, fields.Select(field => field.ToPlainText()));
        Assert.Contains("\\f(a,b)", fields[2].Instruction, StringComparison.Ordinal);
        Assert.Contains("\\r(,x)", fields[3].Instruction, StringComparison.Ordinal);
        Assert.All(fields, field => Assert.All(field.Result.Inlines.OfType<RtfRun>(), run =>
            Assert.Equal(RtfRevisionKind.Inserted, run.RevisionKind)));
    }

    [Fact]
    public void WordToRtf_MapsEquationsInsideHyperlinksAndNestedInlineContentControls() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph._paragraph.Append(new Hyperlink(
            new Run(new Text("link-prefix ")),
            new M.OfficeMath(new M.Run(new M.Text("linked"))),
            new SdtRun(
                new SdtProperties(new SdtId { Val = 2076 }),
                new SdtContentRun(
                    new Run(new Text(" nested-prefix ")),
                    new M.OfficeMath(new M.Run(new M.Text("nested"))),
                    new Run(new Text(" nested-suffix ")))),
            new Run(new Text("link-suffix"))) {
            Anchor = "target"
        });

        RtfParagraph rtfParagraph = Assert.Single(word.ToRtfDocument().Paragraphs);
        RtfField[] fields = rtfParagraph.Inlines.OfType<RtfField>().ToArray();

        Assert.Equal(new[] { "linked", "nested" }, fields.Select(field => field.ToPlainText()));
        Assert.Equal("link-prefix linked nested-prefix nested nested-suffix link-suffix", rtfParagraph.ToPlainText());
        Assert.All(fields, field => Assert.True(field.IsEquation));
    }
}
