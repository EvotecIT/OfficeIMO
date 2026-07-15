using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using M = DocumentFormat.OpenXml.Math;
using W14 = DocumentFormat.OpenXml.Office2010.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains equation tests.
    /// </summary>
    public partial class Word {
        [Fact]
        public void Test_CreatingWordDocumentWithEquation() {
            var filePath = Path.Combine(_directoryWithFiles, "CreatedWithEquation.docx");
            using (var document = WordDocument.Create(filePath)) {
                const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara>";
                document.AddEquation(omml);
                document.Save();
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Single(document.Equations);
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithExponentEquation() {
            var filePath = Path.Combine(_directoryWithFiles, "CreatedWithExponentEquation.docx");
            using (var document = WordDocument.Create(filePath)) {
                const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>";
                document.AddEquation(omml);
                document.Save();
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Single(document.Equations);
            }
        }

        [Fact]
        public void Test_CreatingWordDocumentWithIntegralEquation() {
            var filePath = Path.Combine(_directoryWithFiles, "CreatedWithIntegralEquation.docx");
            using (var document = WordDocument.Create(filePath)) {
                const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:int><m:intPr/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:int></m:oMath></m:oMathPara>";
                document.AddEquation(omml);
                document.Save();
            }

            using (var document = WordDocument.Load(filePath)) {
                Assert.Single(document.Equations);
            }
        }

        [Fact]
        public void TextSetterReplacesComplexMathStructureWithSimpleMathText() {
            var filePath = Path.Combine(_directoryWithFiles, "UpdatedComplexEquationText.docx");
            using (var document = WordDocument.Create(filePath)) {
                const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup></m:oMath></m:oMathPara>";
                WordParagraph equation = document.AddParagraph().AddEquation(omml);

                equation.Text = "total";
                document.Save();
            }

            using WordprocessingDocument package = WordprocessingDocument.Open(filePath, false);
            M.OfficeMath math = Assert.Single(package.MainDocumentPart!.Document.Body!.Descendants<M.OfficeMath>());
            M.Text text = Assert.Single(math.Descendants<M.Text>());
            Assert.Equal("total", text.Text);
            Assert.Empty(math.Descendants<M.Superscript>());
        }

        [Fact]
        public void EquationText_PreservesExplicitlyEmptyOneSidedDelimiters() {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath>" +
                "<m:d><m:dPr><m:begChr m:val=\"{\"/><m:endChr m:val=\"\"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d>" +
                "<m:r><m:t>+</m:t></m:r>" +
                "<m:d><m:dPr><m:begChr m:val=\"\"/><m:endChr m:val=\"}\"/></m:dPr><m:e><m:r><m:t>y</m:t></m:r></m:e></m:d>" +
                "</m:oMath></m:oMathPara>";
            using WordDocument document = WordDocument.Create();

            WordParagraph equationParagraph = document.AddParagraph().AddEquation(omml);

            Assert.Equal("{x+y}", equationParagraph.Text);
            Assert.Equal("{x+y}", Assert.Single(document.Equations).Text);
            Assert.Single(document.Find("{x+y}"));
            Assert.Equal(1, document.FindAndReplace("{x+y}", "z"));
            Assert.Equal("z", equationParagraph.Text);
            Assert.Equal("z", Assert.Single(document.Equations).Text);
        }

        [Fact]
        public void EquationText_UsesDefaultParenthesesOnlyWhenDelimiterPropertiesAreMissing() {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:d><m:dPr/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d></m:oMath></m:oMathPara>";
            using WordDocument document = WordDocument.Create();

            WordParagraph equationParagraph = document.AddParagraph().AddEquation(omml);

            Assert.Equal("(x)", equationParagraph.Text);
        }

        [Fact]
        public void EquationText_DoesNotBorrowDelimiterCharactersFromNestedExpressions() {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath>" +
                "<m:d><m:dPr/><m:e><m:d><m:dPr><m:begChr m:val=\"[\"/><m:endChr m:val=\"]\"/></m:dPr>" +
                "<m:e><m:r><m:t>x</m:t></m:r></m:e></m:d></m:e></m:d>" +
                "</m:oMath></m:oMathPara>";
            using WordDocument document = WordDocument.Create();

            WordParagraph equationParagraph = document.AddParagraph().AddEquation(omml);

            Assert.Equal("([x])", equationParagraph.Text);
        }

        [Fact]
        public void Equation_ProjectsOmmlToLatexMathMlAndLegacyEqField() {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath></m:oMathPara>";
            using WordDocument document = WordDocument.Create();
            document.AddEquation(omml);

            WordEquation equation = Assert.Single(document.Equations);

            Assert.Equal(WordEquationRepresentation.Omml, equation.Representation);
            Assert.Equal("(a)/(b)", equation.Text);
            Assert.Equal("\\frac{a}{b}", equation.ToLatex());
            Assert.Contains("<mfrac>", equation.ToMathMl(), StringComparison.Ordinal);
            Assert.Contains("\\f(a,b)", equation.ToEquationFieldInstruction(), StringComparison.Ordinal);
        }

        [Fact]
        public void Equation_ProjectsEveryOmmlFractionTypeAcrossRepresentations() {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath>" +
                "<m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>" +
                "<m:f><m:fPr><m:type m:val=\"lin\"/></m:fPr><m:num><m:r><m:t>c</m:t></m:r></m:num><m:den><m:r><m:t>d</m:t></m:r></m:den></m:f>" +
                "<m:f><m:fPr><m:type m:val=\"noBar\"/></m:fPr><m:num><m:r><m:t>e</m:t></m:r></m:num><m:den><m:r><m:t>f</m:t></m:r></m:den></m:f>" +
                "<m:f><m:fPr><m:type m:val=\"skw\"/></m:fPr><m:num><m:r><m:t>g</m:t></m:r></m:num><m:den><m:r><m:t>h</m:t></m:r></m:den></m:f>" +
                "</m:oMath></m:oMathPara>";
            using WordDocument document = WordDocument.Create();
            document.AddEquation(omml);

            WordEquation equation = Assert.Single(document.Equations);

            Assert.Equal("(a)/(b)c/dstack(e,f)g⁄h", equation.Text);
            string latex = equation.ToLatex();
            Assert.Contains("\\frac{a}{b}", latex, StringComparison.Ordinal);
            Assert.Contains("{c}/{d}", latex, StringComparison.Ordinal);
            Assert.Contains("\\genfrac{}{}{0pt}{}{e}{f}", latex, StringComparison.Ordinal);
            Assert.Contains("{}^{g}\\! / \\!{}_{h}", latex, StringComparison.Ordinal);

            string mathMl = equation.ToMathMl();
            Assert.Contains("<mfrac><mtext>a</mtext><mtext>b</mtext></mfrac>", mathMl, StringComparison.Ordinal);
            Assert.Contains("<mrow><mtext>c</mtext><mo>/</mo><mtext>d</mtext></mrow>", mathMl, StringComparison.Ordinal);
            Assert.Contains("<mfrac linethickness=\"0\"><mtext>e</mtext><mtext>f</mtext></mfrac>", mathMl, StringComparison.Ordinal);
            Assert.Contains("<mfrac bevelled=\"true\"><mtext>g</mtext><mtext>h</mtext></mfrac>", mathMl, StringComparison.Ordinal);

            string field = equation.ToEquationFieldInstruction();
            Assert.Contains("\\f(a,b)", field, StringComparison.Ordinal);
            Assert.Contains("c/d", field, StringComparison.Ordinal);
            Assert.Contains("\\a\\co1(e,f)", field, StringComparison.Ordinal);
            Assert.Contains("g⁄h", field, StringComparison.Ordinal);
        }

        [Fact]
        public void Equation_ProjectsCommonOmmlStructuresToNativeEqSwitches() {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath>" +
                "<m:rad><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>" +
                "<m:sSup><m:e><m:r><m:t>y</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>" +
                "<m:bar><m:barPr><m:pos m:val=\"top\"/></m:barPr><m:e><m:r><m:t>z</m:t></m:r></m:e></m:bar>" +
                "<m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr></m:m>" +
                "</m:oMath></m:oMathPara>";
            using WordDocument document = WordDocument.Create();
            document.AddEquation(omml);

            string field = Assert.Single(document.Equations).ToEquationFieldInstruction();

            Assert.Contains("\\r(,x)", field, StringComparison.Ordinal);
            Assert.Contains("y\\s\\up8(2)", field, StringComparison.Ordinal);
            Assert.Contains("\\x\\to(z)", field, StringComparison.Ordinal);
            Assert.Contains("\\a\\co2(a,b)", field, StringComparison.Ordinal);
        }

        [Fact]
        public void Equation_ProjectionsEscapeClosingEqParenthesesAndHonorBottomBars() {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath>" +
                "<m:f><m:num><m:r><m:t>a)</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>" +
                "<m:bar><m:barPr><m:pos m:val=\"bot\"/></m:barPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:bar>" +
                "<m:groupChr><m:groupChrPr><m:chr m:val=\"⏟\"/></m:groupChrPr><m:e><m:r><m:t>y</m:t></m:r></m:e></m:groupChr>" +
                "</m:oMath></m:oMathPara>";
            using WordDocument document = WordDocument.Create();
            document.AddEquation(omml);

            WordEquation equation = Assert.Single(document.Equations);

            Assert.Contains("\\f(a\\),b)", equation.ToEquationFieldInstruction(), StringComparison.Ordinal);
            Assert.Contains("\\underline{x}", equation.ToLatex(), StringComparison.Ordinal);
            string mathMl = equation.ToMathMl();
            Assert.Contains("<munder accentunder=\"true\"><mtext>x</mtext><mo>¯</mo></munder>", mathMl, StringComparison.Ordinal);
            Assert.Contains("<munder accentunder=\"true\"><mtext>y</mtext><mo>⏟</mo></munder>", mathMl, StringComparison.Ordinal);
        }

        [Fact]
        public void Equation_ProjectionsApplyOmmlCharacterDefaultsAndDelimiterSeparators() {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath>" +
                "<m:acc><m:e><m:r><m:t>x</m:t></m:r></m:e></m:acc>" +
                "<m:groupChr><m:e><m:r><m:t>y</m:t></m:r></m:e></m:groupChr>" +
                "<m:nary><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>n</m:t></m:r></m:sup><m:e><m:r><m:t>z</m:t></m:r></m:e></m:nary>" +
                "<m:d><m:dPr><m:sepChr m:val=\"|\"/></m:dPr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:d>" +
                "<m:d><m:dPr/><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:d>" +
                "<m:d><m:dPr><m:sepChr m:val=\"\"/></m:dPr><m:e><m:r><m:t>e</m:t></m:r></m:e><m:e><m:r><m:t>f</m:t></m:r></m:e></m:d>" +
                "</m:oMath></m:oMathPara>";
            using WordDocument document = WordDocument.Create();
            document.AddEquation(omml);

            WordEquation equation = Assert.Single(document.Equations);

            Assert.Equal("hat(x)underbrace(y)int_(i)^(n)(z)(a|b)(c│d)(ef)", equation.Text);
            string latex = equation.ToLatex();
            Assert.Contains("\\hat{x}", latex, StringComparison.Ordinal);
            Assert.Contains("\\underbrace{y}", latex, StringComparison.Ordinal);
            Assert.Contains("\\int_{i}^{n} z", latex, StringComparison.Ordinal);
            Assert.Contains("\\left(a\\middle|b\\right)", latex, StringComparison.Ordinal);
            Assert.Contains("\\left(c\\middle|d\\right)", latex, StringComparison.Ordinal);
            Assert.Contains("\\left(ef\\right)", latex, StringComparison.Ordinal);

            string mathMl = equation.ToMathMl();
            Assert.Contains("<mover accent=\"true\"><mtext>x</mtext><mo>̂</mo></mover>", mathMl, StringComparison.Ordinal);
            Assert.Contains("<munder accentunder=\"true\"><mtext>y</mtext><mo>⏟</mo></munder>", mathMl, StringComparison.Ordinal);
            Assert.Contains("<mo>∫</mo>", mathMl, StringComparison.Ordinal);
            Assert.Contains("<mo>|</mo>", mathMl, StringComparison.Ordinal);
            Assert.Contains("<mo>│</mo>", mathMl, StringComparison.Ordinal);

            string field = equation.ToEquationFieldInstruction();
            Assert.Contains("\\o(x,̂)", field, StringComparison.Ordinal);
            Assert.Contains("\\i(i,n,z)", field, StringComparison.Ordinal);
            Assert.Contains("(a|b)", field, StringComparison.Ordinal);
            Assert.Contains("(c│d)", field, StringComparison.Ordinal);
            Assert.Contains("(ef)", field, StringComparison.Ordinal);
        }

        [Fact]
        public void Equation_EscapesFunctionParenthesesInsideEqArguments() {
            const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath>" +
                "<m:f><m:num><m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName><m:e><m:r><m:t>x</m:t></m:r></m:e></m:func></m:num>" +
                "<m:den><m:r><m:t>b</m:t></m:r></m:den></m:f>" +
                "</m:oMath></m:oMathPara>";
            using WordDocument document = WordDocument.Create();
            document.AddEquation(omml);

            string instruction = Assert.Single(document.Equations).ToEquationFieldInstruction();

            Assert.Contains("\\f(sin\\(x\\),b)", instruction, StringComparison.Ordinal);
        }

        [Fact]
        public void EquationOccurrences_DiscoverMathInsideVisibleRevisionWrappers() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("before ");
            var insertedRun = new InsertedRun(new M.OfficeMath(new M.Run(new M.Text("inserted")))) {
                Id = "1",
                Author = "Reviewer"
            };
            var moveToRun = new MoveToRun(new M.OfficeMath(new M.Run(new M.Text("moved")))) {
                Id = "2",
                Author = "Reviewer"
            };
            paragraph._paragraph.Append(insertedRun, moveToRun);
            paragraph.AddText(" after");

            IReadOnlyList<WordEquationOccurrence> occurrences = WordEquation.GetOccurrences(document, paragraph._paragraph);
            List<OpenXmlElement> paragraphChildren = paragraph._paragraph.ChildElements.ToList();

            Assert.Equal(new[] { "inserted", "moved" }, occurrences.Select(occurrence => occurrence.Equation.Text));
            Assert.Equal(
                new[] { paragraphChildren.IndexOf(insertedRun), paragraphChildren.IndexOf(moveToRun) },
                occurrences.Select(occurrence => occurrence.StartChildIndex));
            Assert.Empty(document.ValidateDocument());
        }

        [Fact]
        public void EquationOccurrences_DiscoverOmmlAndEqFieldsInsideInlineContentControls() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("before ");
            var content = new SdtContentRun(
                new M.OfficeMath(new M.Run(new M.Text("omml"))),
                new SimpleField(new Run(new Text("simple"))) { Instruction = " EQ x " },
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode(" EQ \\f(a,b) ")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text("complex")),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
            var contentControl = new SdtRun(
                new SdtProperties(new SdtId { Val = 2076 }),
                content);
            paragraph._paragraph.Append(contentControl);
            paragraph.AddText(" after");

            IReadOnlyList<WordEquationOccurrence> occurrences = WordEquation.GetOccurrences(document, paragraph._paragraph);
            int contentControlIndex = paragraph._paragraph.ChildElements.ToList().IndexOf(contentControl);

            Assert.Equal(new[] { "omml", "simple", "complex" }, occurrences.Select(occurrence => occurrence.Equation.Text));
            Assert.All(occurrences, occurrence => Assert.Equal(contentControlIndex, occurrence.StartChildIndex));
            Assert.Equal(new[] { "omml", "simple", "complex" }, document.Equations.Select(equation => equation.Text));
            Assert.Empty(document.ValidateDocument());
        }

        [Fact]
        public void EquationOccurrences_DiscoverAndOrderMathInsideHyperlinks() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph("before ");
            var hyperlink = new Hyperlink(
                new Run(new Text("link-prefix ")),
                new M.OfficeMath(new M.Run(new M.Text("linked"))),
                new Run(new Text(" link-suffix"))) {
                Anchor = "equation-target"
            };
            paragraph._paragraph.Append(hyperlink);
            paragraph.AddText(" after");

            IReadOnlyList<WordEquationOccurrence> occurrences = WordEquation.GetOccurrences(document, paragraph._paragraph);
            WordEquationOccurrence occurrence = Assert.Single(occurrences);
            IReadOnlyList<WordEquationContentSegment> segments =
                WordEquation.GetVisibleContentSegments(hyperlink, occurrences);

            Assert.Equal("linked", occurrence.Equation.Text);
            Assert.Collection(
                segments,
                segment => {
                    Assert.Equal("link-prefix ", segment.Text);
                    Assert.Same(hyperlink.Elements<Run>().First(), segment.SourceRun);
                },
                segment => Assert.Same(occurrence.Equation, segment.Equation),
                segment => {
                    Assert.Equal(" link-suffix", segment.Text);
                    Assert.Same(hyperlink.Elements<Run>().Last(), segment.SourceRun);
                });
            Assert.Equal("link-prefix linked link-suffix", WordEquation.GetVisibleTextWithEquations(hyperlink, occurrences));
            WordParagraph equationParagraph = Assert.Single(document.ParagraphsEquations);
            Assert.True(equationParagraph.IsHyperLink);
            Assert.Equal("linked", equationParagraph.Equation!.Text);
            Assert.Empty(document.ValidateDocument());
        }

        [Fact]
        public void EquationContentSegments_PreserveHyperlinkContextAndOrderedRunArtifacts() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();
            var breakRun = new Run(new Text("before-break"), new Break(), new Text("after-break"));
            var hyperlink = new Hyperlink(
                new Run(new Text("prefix")),
                new M.OfficeMath(new M.Run(new M.Text("linked"))),
                breakRun,
                new Run(new Text("suffix"))) {
                Anchor = "equation-target"
            };
            paragraph._paragraph.Append(hyperlink);

            IReadOnlyList<WordEquationOccurrence> occurrences = WordEquation.GetOccurrences(document, paragraph._paragraph);
            IReadOnlyList<WordEquationContentSegment> segments = WordEquation.GetVisibleContentSegments(hyperlink, occurrences);

            Assert.Collection(
                segments,
                segment => Assert.Equal("prefix", segment.Text),
                segment => {
                    Assert.Equal("linked", segment.Equation?.Text);
                    Assert.Same(hyperlink, segment.SourceElement);
                    Assert.True(segment.CreateSourceParagraph(document, paragraph._paragraph, paragraph).IsHyperLink);
                },
                segment => Assert.Equal("before-break", segment.Text),
                segment => {
                    Assert.True(segment.IsRunArtifact);
                    Assert.Same(breakRun, segment.SourceRun);
                    Assert.IsType<Break>(segment.ArtifactElement);
                },
                segment => Assert.Equal("after-break", segment.Text),
                segment => Assert.Equal("suffix", segment.Text));
            Assert.Equal("prefixlinkedbefore-break\nafter-breaksuffix", WordEquation.GetVisibleTextWithEquations(hyperlink, occurrences));
        }

        [Fact]
        public void EquationContentSegments_PreserveHyperlinkContextThroughInlineContentControl() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();
            var contentControl = new SdtRun(
                new SdtProperties(new SdtId { Val = 2078 }),
                new SdtContentRun(
                    new Run(new Text("prefix")),
                    new M.OfficeMath(new M.Run(new M.Text("nested-linked"))),
                    new Run(new Text("suffix"))));
            var hyperlink = new Hyperlink(contentControl) { Anchor = "nested-equation-target" };
            paragraph._paragraph.Append(hyperlink);

            WordEquationOccurrence occurrence = Assert.Single(WordEquation.GetOccurrences(document, paragraph._paragraph));
            WordEquationContentSegment equationSegment = Assert.Single(
                WordEquation.GetVisibleContentSegments(hyperlink, new[] { occurrence }),
                segment => segment.Equation != null);

            Assert.Same(contentControl, equationSegment.SourceElement);
            WordParagraph source = equationSegment.CreateSourceParagraph(document, paragraph._paragraph, paragraph);
            Assert.True(source.IsStructuredDocumentTag);
            Assert.True(source.IsHyperLink);
            Assert.Equal("nested-equation-target", source.Hyperlink?.Anchor);
        }

        [Fact]
        public void EquationContentSegments_FormControlVisibleTextDoesNotDuplicateNestedEquation() {
            using WordDocument document = WordDocument.Create();
            WordParagraph paragraph = document.AddParagraph();
            var contentControl = new SdtRun(
                new SdtProperties(
                    new W14.SdtContentCheckBox(new W14.Checked { Val = W14.OnOffValues.One })),
                new SdtContentRun(
                    new Run(new Text("☑")),
                    new M.OfficeMath(new M.Run(new M.Text("approved")))));
            paragraph._paragraph.Append(contentControl);

            IReadOnlyList<WordEquationOccurrence> occurrences = WordEquation.GetOccurrences(document, paragraph._paragraph);
            IReadOnlyList<WordEquationContentSegment> segments = WordEquation.GetVisibleContentSegments(contentControl, occurrences);

            Assert.Collection(
                segments,
                segment => {
                    Assert.True(segment.IsRunArtifact);
                    Assert.Equal("☑", segment.ArtifactVisibleText);
                },
                segment => Assert.Equal("approved", segment.Equation?.Text));
            Assert.Equal("☑approved", WordEquation.GetVisibleTextWithEquations(contentControl, occurrences));
        }

        [Fact]
        public void WordFieldType_ExistingNumericValuesRemainStableWhenEqIsAdded() {
            Assert.Equal(19, (int)WordFieldType.FileName);
            Assert.Equal(71, (int)WordFieldType.Formula);
            Assert.Equal(72, (int)WordFieldType.EQ);
        }
    }
}
