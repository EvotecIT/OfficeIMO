using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using M = DocumentFormat.OpenXml.Math;
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
                "</m:oMath></m:oMathPara>";
            using WordDocument document = WordDocument.Create();
            document.AddEquation(omml);

            WordEquation equation = Assert.Single(document.Equations);

            Assert.Contains("\\f(a\\),b)", equation.ToEquationFieldInstruction(), StringComparison.Ordinal);
            Assert.Contains("\\underline{x}", equation.ToLatex(), StringComparison.Ordinal);
        }

        [Fact]
        public void WordFieldType_ExistingNumericValuesRemainStableWhenEqIsAdded() {
            Assert.Equal(19, (int)WordFieldType.FileName);
            Assert.Equal(71, (int)WordFieldType.Formula);
            Assert.Equal(72, (int)WordFieldType.EQ);
        }
    }
}
