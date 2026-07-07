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
                document.Save(false);
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
                document.Save(false);
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
                document.Save(false);
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
                document.Save(false);
            }

            using WordprocessingDocument package = WordprocessingDocument.Open(filePath, false);
            M.OfficeMath math = Assert.Single(package.MainDocumentPart!.Document.Body!.Descendants<M.OfficeMath>());
            M.Text text = Assert.Single(math.Descendants<M.Text>());
            Assert.Equal("total", text.Text);
            Assert.Empty(math.Descendants<M.Superscript>());
        }
    }
}
