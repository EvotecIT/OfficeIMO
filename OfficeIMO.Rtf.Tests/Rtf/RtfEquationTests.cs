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
        RtfField hyperlinkField = Assert.Single(rtfParagraph.Inlines.OfType<RtfField>());
        RtfField[] fields = hyperlinkField.Result.Inlines.OfType<RtfField>().ToArray();

        Assert.NotNull(hyperlinkField.HyperlinkField);
        Assert.Equal("target", hyperlinkField.HyperlinkField!.SubAddress);
        Assert.Equal(new[] { "linked", "nested" }, fields.Select(field => field.ToPlainText()));
        Assert.Equal("link-prefix linked nested-prefix nested nested-suffix link-suffix", rtfParagraph.ToPlainText());
        Assert.All(fields, field => Assert.True(field.IsEquation));

        string serialized = word.ToRtfDocument().ToRtf();
        RtfField parsedHyperlink = Assert.Single(Assert.Single(RtfDocument.Read(serialized).Document.Paragraphs).Inlines.OfType<RtfField>());
        Assert.Equal(2, parsedHyperlink.Result.Inlines.OfType<RtfField>().Count(field => field.IsEquation));

        using WordDocument roundTrip = RtfDocument.Read(serialized).Document.ToWordDocument();
        Hyperlink roundTripHyperlink = Assert.Single(Assert.Single(roundTrip.Paragraphs)._paragraph.Elements<Hyperlink>());
        Assert.Equal(2, roundTripHyperlink.Elements<SimpleField>().Count());
        Assert.Equal(2, roundTrip.Equations.Count);
        var hyperlinkErrors = roundTrip.ValidateDocument()
            .Where(error => error.Node is Hyperlink || error.Node?.Ancestors<Hyperlink>().Any() == true)
            .ToArray();
        Assert.True(
            hyperlinkErrors.Length == 0,
            string.Join(Environment.NewLine, hyperlinkErrors.Select(error =>
                $"{error.Description}{Environment.NewLine}{error.Node?.OuterXml}")));

        string docPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".doc");
        try {
            roundTrip.Save(docPath);
            using WordDocument docReload = WordDocument.Load(docPath);
            Hyperlink docHyperlink = Assert.Single(Assert.Single(docReload.Paragraphs)._paragraph.Elements<Hyperlink>());
            Assert.Equal(2, docHyperlink.Elements<SimpleField>().Count());
            Assert.Equal(new[] { "linked", "nested" }, docReload.Equations.Select(equation => equation.Text));
            var docErrors = docReload.ValidateDocument().ToArray();
            Assert.True(
                docErrors.Length == 0,
                string.Join(Environment.NewLine, docErrors.Select(error =>
                    $"{error.Description}{Environment.NewLine}{error.Node?.OuterXml}")));
        } finally {
            if (File.Exists(docPath)) File.Delete(docPath);
        }
    }

    [Fact]
    public void WordToRtf_KeepsEquationOnlyHyperlinkAsNestedEqField() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph._paragraph.Append(new Hyperlink(
            new M.OfficeMath(new M.Run(new M.Text("linked-only")))) {
            Anchor = "target"
        });

        RtfField hyperlinkField = Assert.Single(Assert.Single(word.ToRtfDocument().Paragraphs).Inlines.OfType<RtfField>());

        Assert.Equal("target", hyperlinkField.HyperlinkField?.SubAddress);
        RtfField equationField = Assert.Single(hyperlinkField.Result.Inlines.OfType<RtfField>());
        Assert.True(equationField.IsEquation);
        Assert.Equal("linked-only", equationField.ToPlainText());
    }

    [Fact]
    public void WordRtfRoundTrip_KeepsComplexEqFieldInsideHyperlink() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph._paragraph.Append(new Hyperlink(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" EQ \\f(a,b) ")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("(a)/(b)")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End })) {
            Anchor = "target"
        });

        RtfDocument rtf = word.ToRtfDocument();
        RtfField hyperlinkField = Assert.Single(Assert.Single(rtf.Paragraphs).Inlines.OfType<RtfField>());
        RtfField equationField = Assert.Single(hyperlinkField.Result.Inlines.OfType<RtfField>());
        Assert.True(equationField.IsEquation);
        Assert.Contains("\\f(a,b)", equationField.Instruction, StringComparison.Ordinal);

        using WordDocument roundTrip = RtfDocument.Read(rtf.ToRtf()).Document.ToWordDocument();
        Hyperlink roundTripHyperlink = Assert.Single(Assert.Single(roundTrip.Paragraphs)._paragraph.Elements<Hyperlink>());
        Assert.Single(roundTripHyperlink.Elements<SimpleField>());
        WordEquation equation = Assert.Single(roundTrip.Equations);
        Assert.Equal(WordEquationRepresentation.EquationField, equation.Representation);
        Assert.Contains("\\f(a,b)", equation.FieldInstruction!, StringComparison.Ordinal);
        Assert.DoesNotContain(roundTrip.ValidateDocument(), error =>
            error.Node is Hyperlink || error.Node?.Ancestors<Hyperlink>().Any() == true);
    }

    [Fact]
    public void WordToRtf_CapturesOmmlInsideActiveComplexHyperlinkResult() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph._paragraph.Append(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" HYPERLINK \\l \"target\" ")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("prefix ")),
            new M.OfficeMath(new M.Run(new M.Text("captured-equation"))),
            new SimpleField(new Run(new Text("simple-equation"))) {
                Instruction = " EQ \\f(s,t) "
            },
            new SdtRun(
                new SdtProperties(new SdtId { Val = 2080 }),
                new SdtContentRun(
                    new Run(new Text(" controlled-prefix ")),
                    new M.OfficeMath(new M.Run(new M.Text("controlled-equation"))))),
            new InsertedRun(
                new Run(new Text(" revised-prefix ")),
                new M.OfficeMath(new M.Run(new M.Text("revised-equation")))) {
                Id = "2081",
                Author = "Reviewer"
            },
            new Run(new Text(" suffix")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        RtfParagraph rtfParagraph = Assert.Single(word.ToRtfDocument().Paragraphs);
        RtfField hyperlinkField = Assert.Single(rtfParagraph.Inlines.OfType<RtfField>());
        RtfField[] equationFields = hyperlinkField.Result.Inlines.OfType<RtfField>().ToArray();

        Assert.NotNull(hyperlinkField.HyperlinkField);
        Assert.Equal("target", hyperlinkField.HyperlinkField!.SubAddress);
        Assert.All(equationFields, field => Assert.True(field.IsEquation));
        Assert.Equal(
            new[] { "captured-equation", "simple-equation", "controlled-equation", "revised-equation" },
            equationFields.Select(field => field.ToPlainText()));
        Assert.Equal(
            "prefix captured-equationsimple-equation controlled-prefix controlled-equation revised-prefix revised-equation suffix",
            hyperlinkField.ToPlainText());
        Assert.All(equationFields[3].Result.Inlines.OfType<RtfRun>(), run =>
            Assert.Equal(RtfRevisionKind.Inserted, run.RevisionKind));
        Assert.Empty(rtfParagraph.Inlines.Skip(1));
    }

    [Fact]
    public void WordToRtf_PreservesRevisionWrappedContainersInsideActiveComplexFieldResult() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph._paragraph.Append(
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode(" HYPERLINK \\l \"outer\" ")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new InsertedRun(
                new Hyperlink(
                    new Run(new Text("linked-prefix ")),
                    new M.OfficeMath(new M.Run(new M.Text("linked-equation"))),
                    new Run(new Text(" linked-suffix"))) {
                    Anchor = "inner"
                },
                new SdtRun(
                    new SdtProperties(new SdtId { Val = 2082 }),
                    new SdtContentRun(
                        new Run(new Text(" controlled-prefix ")),
                        new M.OfficeMath(new M.Run(new M.Text("controlled-equation"))),
                        new Run(new Text(" controlled-suffix"))))) {
                Id = "2083",
                Author = "Reviewer"
            },
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        RtfParagraph rtfParagraph = Assert.Single(word.ToRtfDocument().Paragraphs);
        RtfField outerHyperlink = Assert.Single(rtfParagraph.Inlines.OfType<RtfField>());
        RtfField innerHyperlink = Assert.Single(
            outerHyperlink.Result.Inlines.OfType<RtfField>(),
            field => field.HyperlinkField?.SubAddress == "inner");
        RtfField linkedEquation = Assert.Single(innerHyperlink.Result.Inlines.OfType<RtfField>());
        RtfField controlledEquation = Assert.Single(
            outerHyperlink.Result.Inlines.OfType<RtfField>(),
            field => field.IsEquation);

        Assert.Equal("outer", outerHyperlink.HyperlinkField?.SubAddress);
        Assert.Equal("linked-equation", linkedEquation.ToPlainText());
        Assert.Equal("controlled-equation", controlledEquation.ToPlainText());
        Assert.Equal(
            "linked-prefix linked-equation linked-suffix controlled-prefix controlled-equation controlled-suffix",
            outerHyperlink.ToPlainText());
        Assert.All(innerHyperlink.Result.Inlines.OfType<RtfRun>(), run =>
            Assert.Equal(RtfRevisionKind.Inserted, run.RevisionKind));
        Assert.All(linkedEquation.Result.Inlines.OfType<RtfRun>(), run =>
            Assert.Equal(RtfRevisionKind.Inserted, run.RevisionKind));
        Assert.All(controlledEquation.Result.Inlines.OfType<RtfRun>(), run =>
            Assert.Equal(RtfRevisionKind.Inserted, run.RevisionKind));
        Assert.Empty(rtfParagraph.Inlines.Skip(1));
    }

    [Fact]
    public void WordToRtf_FlattensRevisedActiveFieldsButKeepsRevisedEquationFields() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph();
        paragraph._paragraph.Append(
            new InsertedRun(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" INCLUDEPICTURE \\\"https://example.test/active.png\\\" ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("safe cached image")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End })) {
                Id = "3001",
                Author = "Reviewer",
            },
            new InsertedRun(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" EQ \\f(a,b) ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("(a)/(b)")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End })) {
                Id = "3002",
                Author = "Reviewer",
            },
            new InsertedRun(
                new SimpleField(new Run(new Text("safe simple result"))) {
                    Instruction = " INCLUDEPICTURE \\\"https://example.test/simple.png\\\" ",
                }) {
                Id = "3003",
                Author = "Reviewer",
            },
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new InsertedRun(new Run(new FieldCode(" INCLUDEPICTURE \\\"https://example.test/split.png\\\" "))) {
                Id = "3004",
                Author = "Reviewer",
            },
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text("safe split result")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End }));

        RtfParagraph rtfParagraph = Assert.Single(word.ToRtfDocument().Paragraphs);
        RtfField field = Assert.Single(rtfParagraph.Inlines.OfType<RtfField>());

        Assert.True(field.IsEquation);
        Assert.Equal("(a)/(b)", field.ToPlainText());
        Assert.Equal("safe cached image(a)/(b)safe simple resultsafe split result", rtfParagraph.ToPlainText());
        string serialized = word.ToRtfDocument().ToRtf();
        Assert.DoesNotContain("INCLUDEPICTURE", serialized, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("example.test", serialized, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void WordToRtf_MapsTopLevelInlineContentControlWithoutDroppingNestedEquation() {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("before ");
        paragraph._paragraph.Append(new SdtRun(
            new SdtProperties(new SdtId { Val = 2076 }),
            new SdtContentRun(
                new Run(new Text("control-prefix ")),
                new M.OfficeMath(new M.Run(new M.Text("controlled"))),
                new Run(new Text(" control-suffix")))));
        paragraph.AddText(" after");

        RtfParagraph rtfParagraph = Assert.Single(word.ToRtfDocument().Paragraphs);

        Assert.Equal("before control-prefix controlled control-suffix after", rtfParagraph.ToPlainText());
        RtfField equationField = Assert.Single(rtfParagraph.Inlines.OfType<RtfField>());
        Assert.True(equationField.IsEquation);
        Assert.Equal("controlled", equationField.ToPlainText());
    }
}
