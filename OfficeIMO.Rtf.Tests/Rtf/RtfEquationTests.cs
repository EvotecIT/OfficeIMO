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
            });

        RtfDocument rtf = word.ToRtfDocument();
        RtfField[] fields = Assert.Single(rtf.Paragraphs).Inlines.OfType<RtfField>().ToArray();

        Assert.Equal(new[] { "inserted", "moved" }, fields.Select(field => field.ToPlainText()));
        Assert.All(fields, field => Assert.All(field.Result.Inlines.OfType<RtfRun>(), run =>
            Assert.Equal(RtfRevisionKind.Inserted, run.RevisionKind)));
    }
}
